<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
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
  exclude-result-prefixes="odf style text number draw page"> 
  <xsl:template name ="Notes" match ="/office:document-content/office:body/office:presentation/draw:page/presentation:notes" mode="Notes">
    <xsl:param name ="pageNo"/>
    <p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
			   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1" name=""/>
            <p:cNvGrpSpPr/>
            <p:nvPr/>
          </p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="0" cy="0"/>
              <a:chOff x="0" y="0"/>
              <a:chExt cx="0" cy="0"/>
            </a:xfrm>
          </p:grpSpPr>
          <xsl:for-each select ="./draw:page-thumbnail">
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr name="Slide Image Placeholder 1">
                  <xsl:attribute name="id">
                    <xsl:value-of select="position()+1"/>
                  </xsl:attribute>
                </p:cNvPr>
                <p:cNvSpPr>
                  <a:spLocks noGrp="1"/>
                </p:cNvSpPr>
                <p:nvPr>
                    <p:ph type="sldImg"/>
                  </p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <xsl:if test="@svg:width and @svg:height and @svg:x and @svg:y">
                  <!--Write the cordinates-->
                  <xsl:call-template name="tmpgetCordinates"/>
                </xsl:if>
                <a:ln />
              </p:spPr>
            </p:sp>
          </xsl:for-each>
          <xsl:variable name ="flagShapeNotes">
            <xsl:for-each select ="draw:frame">
              <xsl:if test ="not(./@presentation:class[contains(.,'notes')])">
                <xsl:value-of  select ="'true'"/>
              </xsl:if>              
            </xsl:for-each>
          </xsl:variable>
          <xsl:if test ="$flagShapeNotes!=''">
            <!-- warn if shapes in notes -->
            <xsl:message terminate="no">translation.odf2oox.notesTypeShapesInNotes</xsl:message>
          </xsl:if>
          <xsl:for-each select ="draw:frame[@presentation:class[contains(.,'notes')]]">
            <xsl:variable name ="masterPageName" select ="./parent::node()/@draw:master-page-name"/>
            <xsl:variable name="FrameCount" select="concat('Frame',position())"/>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr name="Notes Placeholder 2">
                    <xsl:attribute name ="id">
                      <xsl:value-of select ="position() + 2"/>
                    </xsl:attribute>
                  </p:cNvPr>
                  <p:cNvSpPr>
                    <a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>
                  </p:cNvSpPr>
                  <p:nvPr>
                    <p:ph type="body" idx="1"/>
                  </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                  <xsl:if test="@svg:width and @svg:height and @svg:x and @svg:y">
                    <!--Write the cordinates-->
                    <xsl:call-template name="tmpgetCordinates"/>
                  </xsl:if>
                  <!-- Solid fill color -->
                  <xsl:call-template name ="fillColor" >
                    <xsl:with-param name ="prId" select ="@presentation:style-name" />
                  </xsl:call-template>
                  <!-- added by Vipul to insert line style-->
                  <!--start-->
                  <xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
                  <xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
                    <xsl:call-template name ="getFillColor"/>
                    <xsl:call-template name ="getLineStyle"/>
                  </xsl:for-each>
                  <!--End-->
                </p:spPr>
                <p:txBody>
                  <xsl:call-template name ="tmpNotesTextAlignment" >
                    <xsl:with-param name ="prId" select ="@presentation:style-name"/>
                  </xsl:call-template >
                  <a:lstStyle/>
                  <xsl:call-template name ="tmpNotesprocessText" >
                    <xsl:with-param name ="layoutName" select ="@presentation:class"/>
                    <xsl:with-param name ="FrameCount" select ="$FrameCount"/>
                    <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template >
                </p:txBody>
              </p:sp>
          </xsl:for-each >
          <!-- Code for footer , slide number and date time control -->
          <xsl:variable name ="pageStyle">
            <xsl:value-of select ="@draw:style-name"/>
          </xsl:variable>
          <xsl:variable name ="tmpNotesfooterId">
            <xsl:value-of select ="@presentation:use-footer-name"/>
          </xsl:variable>
          <xsl:variable name ="dateId">
            <xsl:value-of select ="@presentation:use-date-time-name"/>
          </xsl:variable>
          <xsl:for-each select ="document('content.xml')//style:style[@style:name=$pageStyle]">
            <xsl:if test ="style:drawing-page-properties[@presentation:display-footer='true']">
              <xsl:if test ="document('content.xml')//presentation:footer-decl[@presentation:name=$tmpNotesfooterId]">
                <xsl:call-template name ="tmpNotesfooter" >
                  <xsl:with-param name ="tmpNotesfooterId" select ="$tmpNotesfooterId"/>
                </xsl:call-template >
              </xsl:if>
            </xsl:if>
            <xsl:if test ="style:drawing-page-properties[@presentation:display-date-time='true']">
              <xsl:if test ="document('content.xml')//presentation:date-time-decl[@presentation:name=$dateId]">
                <xsl:call-template name ="tmpNotesfooterDate" >
                  <xsl:with-param name ="tmpNotesfooterDateId" select ="$dateId"/>
                </xsl:call-template >
              </xsl:if>
            </xsl:if >
            <xsl:if test ="style:drawing-page-properties[@presentation:display-page-number='true']">
              <xsl:call-template name ="tmpNoteslideNumber">
                <xsl:with-param name ="pageNumber" select ="$pageNo"/>
              </xsl:call-template >
            </xsl:if >
          </xsl:for-each>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:notes>
  </xsl:template>
  <xsl:template name ="tmpNotesprocessText" >
    <xsl:param name ="layoutName"/>
    <xsl:param name ="FrameCount"/>
    <!-- Paremeter added by vijayeta,get master page name, dated:13-7-07-->
    <xsl:param name ="masterPageName"/>
    <xsl:variable name ="defFontSize">
      <xsl:value-of  select ="office:document-styles/office:styles/style:style/style:text-properties/@fo:font-size"/>
    </xsl:variable>
    <!-- Added by lohith.ar - Start - Mouse click events -->
    <xsl:variable name="PostionCount">
      <xsl:value-of select="position()"/>
    </xsl:variable>
    <!-- End - variable to set the hyperlinks values for Text inside Frame -->
    <xsl:variable name ="prClassName">
      <xsl:value-of select ="'outline'"/>
    </xsl:variable>
    <xsl:choose >
      <xsl:when test ="draw:text-box/text:p/text:span">
        <xsl:for-each select ="node()">
          <!--<xsl:for-each select ="draw:text-box/text:p">-->
          <xsl:for-each select ="child::node()[position()]">
            <xsl:choose >
              <xsl:when test ="name()='text:p'">
                <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <xsl:call-template name ="tmpNotesparaProperties" >
                    <xsl:with-param name ="paraId" >
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:with-param >
                    <xsl:with-param name ="isBulleted" select ="'false'"/>
                    <xsl:with-param name ="level" select ="'0'"/>
                    <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
                    <xsl:with-param name ="isNumberingEnabled" select ="'false'"/>
                  </xsl:call-template >
                  <xsl:for-each select ="child::node()[position()]">
                    <xsl:choose >
                      <xsl:when test ="name()='text:span'">
                        <!-- Added by lohith - bug fix 1731885 -->
                        <xsl:if test="node()">
                          <a:r>
                            <a:rPr lang="en-US" smtClean="0">
                              <xsl:variable name ="textId">
                                <xsl:value-of select ="@text:style-name"/>
                              </xsl:variable>
                              <xsl:if test ="not($textId ='')">
                                <xsl:call-template name ="tmpfontStyles">
                                  <xsl:with-param name ="TextStyleID" select ="$textId" />
                                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                                  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                                </xsl:call-template>
                              </xsl:if>
                            </a:rPr >
                            <a:t>
                              <xsl:call-template name ="insertTab" />
                            </a:t>
                          </a:r>
                          <xsl:if test ="text:line-break">
                            <xsl:call-template name ="processBR">
                              <xsl:with-param name ="T" select ="@text:style-name" />
                              <xsl:with-param name ="prClassName" select ="$prClassName"/>
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
                            <xsl:if test ="not($textId ='')">
                              <xsl:call-template name ="tmpfontStyles">
                                <xsl:with-param name ="TextStyleID" select ="$textId" />
                                <xsl:with-param name ="prClassName" select ="$prClassName"/>
                                <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                              </xsl:call-template>
                            </xsl:if>
                          </a:endParaRPr>
                        </xsl:if>

                      </xsl:when >
                      <xsl:when test ="name()='text:line-break'">
                        <xsl:call-template name ="processBR">
                          <xsl:with-param name ="T" select ="@text:style-name" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        </xsl:call-template>
                      </xsl:when>
                      <xsl:when test ="not(name()='text:span')">
                        <a:r>
                          <a:rPr lang="en-US" smtClean="0">
                            <!--Font Size -->
                            <xsl:variable name ="textId">
                              <xsl:value-of select ="@text:style-name"/>
                            </xsl:variable>
                            <xsl:if test ="not($textId ='')">
                              <xsl:call-template name ="tmpfontStyles">
                                <xsl:with-param name ="TextStyleID" select ="$textId" />
                                <xsl:with-param name ="prClassName" select ="$prClassName"/>
                                <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                              </xsl:call-template>
                            </xsl:if>
                          </a:rPr >
                          <a:t>
                            <xsl:call-template name ="insertTab" />
                          </a:t>
                        </a:r>
                      </xsl:when >
                    </xsl:choose>
                  </xsl:for-each>
                </a:p >
              </xsl:when>
              <xsl:when test ="name()='text:list'">
                <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <xsl:variable name ="lvl">
                    <xsl:if test ="text:list-item/text:list">
                      <xsl:call-template name ="getListLevel">
                        <xsl:with-param name ="levelCount"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test ="not(text:list-item/text:list)">
                      <xsl:value-of select ="'0'"/>
                    </xsl:if>
                  </xsl:variable >
                  <xsl:variable name="paragraphId" >
                    <xsl:call-template name ="getParaStyleName">
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:variable name ="isNumberingEnabled">
                    <xsl:if test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                      <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
                    </xsl:if>
                    <xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
                      <xsl:value-of select ="'true'"/>
                    </xsl:if>
                  </xsl:variable>
                  <xsl:call-template name ="tmpNotesparaProperties" >
                    <xsl:with-param name ="paraId" >
                      <xsl:value-of select ="$paragraphId"/>
                    </xsl:with-param >
                    <!-- list property also included-->
                    <xsl:with-param name ="listId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:with-param >
                    <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
                    <xsl:with-param name ="isBulleted" select ="'true'"/>
                    <xsl:with-param name ="level" select ="$lvl"/>
                    <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
                    <!-- parameter added by vijayeta, dated 11-7-07-->
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template >
                  <!--End of Code inserted by Vijayets for Bullets and numbering-->
                  <xsl:for-each select ="child::node()[position()]">
                    <xsl:choose >
                      <xsl:when test ="name()='text:list-item'">
                        <xsl:variable name ="currentNodeStyle">
                          <xsl:call-template name ="getTextNodeForFontStyle">
                            <xsl:with-param name ="prClassName" select ="$prClassName"/>
                            <xsl:with-param name ="lvl" select ="$lvl"/>
                            <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:copy-of select ="$currentNodeStyle"/>
                      </xsl:when >
                    </xsl:choose>
                  </xsl:for-each>
                </a:p >
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:for-each >
      </xsl:when>
      <xsl:when test ="draw:text-box/text:list/text:list-item">
        <!--<xsl:when test ="draw:text-box/text:list/text:list-item/text:p/text:span">-->
        <xsl:for-each select ="draw:text-box/text:list">
          <xsl:variable name="forCount" select="position()" />
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <!--Code inserted by Vijayets for Bullets and numbering-->
            <!--Check if Levels are present-->
            <xsl:variable name ="lvl">
              <xsl:if test ="text:list-item/text:list">
                <xsl:call-template name ="getListLevel">
                  <xsl:with-param name ="levelCount"/>
                </xsl:call-template>
              </xsl:if>
              <xsl:if test ="not(text:list-item/text:list)">
                <xsl:value-of select ="'0'"/>
              </xsl:if>
            </xsl:variable >
            <xsl:variable name="paragraphId" >
              <xsl:call-template name ="getParaStyleName">
                <xsl:with-param name ="lvl" select ="$lvl"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="isNumberingEnabled">
              <xsl:if test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:if>
              <xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:variable>
            <xsl:call-template name ="tmpNotesparaProperties" >
              <xsl:with-param name ="paraId" >
                <xsl:value-of select ="$paragraphId"/>
              </xsl:with-param >
              <!-- list property also included-->
              <xsl:with-param name ="listId">
                <xsl:value-of select ="@text:style-name"/>
              </xsl:with-param >
              <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
              <xsl:with-param name ="isBulleted" select ="'true'"/>
              <xsl:with-param name ="level" select ="$lvl"/>
              <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
              <!-- parameter added by vijayeta, dated 13-7-07-->
              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
            </xsl:call-template >
            <!--End of Code inserted by Vijayets for Bullets and numbering-->
            <xsl:for-each select ="child::node()[position()]">
              <xsl:choose >
                <xsl:when test ="name()='text:list-item'">
                  <xsl:variable name ="currentNodeStyle">
                    <xsl:call-template name ="getTextNodeForFontStyle">
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      <xsl:with-param name ="fileName" select ="'content.xml'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:copy-of select ="$currentNodeStyle"/>
                </xsl:when >
              </xsl:choose>
            </xsl:for-each>
          </a:p >
        </xsl:for-each >
      </xsl:when> 
      <xsl:when test ="draw:text-box/text:list/text:list-item/text:p">
        <xsl:for-each select ="draw:text-box/text:list">
          <a:p  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <!--Check if Levels are present-->
            <xsl:variable name ="lvl">
              <xsl:if test ="text:list-item/text:list">
                <xsl:call-template name ="getListLevel">
                  <xsl:with-param name ="levelCount"/>
                </xsl:call-template>
              </xsl:if>
              <xsl:if test ="not(text:list-item/text:list)">
                <xsl:value-of select ="'0'"/>
              </xsl:if>
            </xsl:variable >
            <xsl:variable name="paragraphId" >
              <xsl:call-template name ="getParaStyleName">
                <xsl:with-param name ="lvl" select ="$lvl"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="isNumberingEnabled">
              <xsl:if test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:if>
              <xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:variable>
            <xsl:call-template name ="tmpNotesparaProperties" >
              <xsl:with-param name ="paraId" >
                <xsl:value-of select ="$paragraphId"/>
              </xsl:with-param >
              <!-- list property also included-->
              <xsl:with-param name ="listId">
                <xsl:value-of select ="@text:style-name"/>
              </xsl:with-param >

              <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
              <xsl:with-param name ="isBulleted" select ="'true'"/>
              <xsl:with-param name ="level" select ="$lvl"/>
              <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
              <!-- parameter added by vijayeta, dated 11-7-07-->
              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
            </xsl:call-template >
            <!--End of Code inserted by Vijayets for Bullets and numbering-->
            <xsl:for-each select ="child::node()[position()]">
              <xsl:choose >
                <xsl:when test ="name()='text:list-item'">
                  <xsl:variable name ="currentNodeStyle">
                    <xsl:call-template name ="getTextNodeForFontStyle">
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <!-- parameter added by vijayeta, dated 11-7-07-->
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      <xsl:with-param name ="fileName" select ="'content.xml'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:copy-of select ="$currentNodeStyle"/>
                </xsl:when >
              </xsl:choose>
            </xsl:for-each>
          </a:p >
        </xsl:for-each >
      </xsl:when>
      <xsl:when test ="draw:text-box/text:p">
        <xsl:for-each select ="draw:text-box/text:p">
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xsl:call-template name ="tmpNotesparaProperties" >
              <xsl:with-param name ="paraId" >
                <xsl:value-of select ="@text:style-name"/>
              </xsl:with-param >
              <xsl:with-param name ="isBulleted" select ="'false'"/>
              <xsl:with-param name ="level" select ="'0'"/>
              <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
              <xsl:with-param name ="isNumberingEnabled" select ="'false'"/>
            </xsl:call-template >
            <a:r >
              <a:rPr lang="en-US" smtClean="0">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="not($textId ='')">
                  <xsl:call-template name ="tmpfontStyles">
                    <xsl:with-param name ="TextStyleID" select ="$textId" />
                    <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template>
                </xsl:if>
              </a:rPr >
              <a:t>
                <xsl:call-template name ="insertTab" />
              </a:t>
            </a:r>
            <xsl:if test ="text:span/text:line-break">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="text:span/@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
          </a:p >
        </xsl:for-each >
      </xsl:when>
      <xsl:otherwise >
        <a:p>
          <a:r >
            <a:rPr lang="en-US" smtClean="0">
              <a:latin charset="0" typeface="Arial" />
            </a:rPr >
            <a:t>
              <xsl:call-template name ="insertTab" />
            </a:t>
          </a:r>
        </a:p >
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name ="tmpNotesfooter">
    <xsl:param name ="tmpNotesfooterId"></xsl:param>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="6" name="footer Placeholder 5" />
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="ftr" sz="quarter" idx="11" />
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr >
        <!-- footer date layout details style:master-page style:name -->
        <xsl:call-template name ="tmpNotesGetFrameDetails">
          <xsl:with-param name ="LayoutName" select ="'footer'"/>
        </xsl:call-template>
      </p:spPr >
      <p:txBody>
        <a:bodyPr />
        <a:lstStyle />
        <a:p>
          <a:r>
            <a:rPr lang="en-US" smtClean="0" />
            <a:t>
              <xsl:for-each select ="document('content.xml') //presentation:footer-decl[@presentation:name=$tmpNotesfooterId] ">
                <xsl:value-of select ="."/>
              </xsl:for-each >
            </a:t>
          </a:r>
          <!--<a:endParaRPr lang="en-US" />-->
        </a:p>
      </p:txBody>
    </p:sp >
  </xsl:template>
  <xsl:template name ="tmpNoteslideNumber">
    <xsl:param name ="pageNumber"/>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="5" name="Slide Number Placeholder 4" />
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="sldNum" sz="quarter" idx="12" />
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr >
        <!-- footer slide number layout details style:master-page style:name -->
        <xsl:call-template name ="tmpNotesGetFrameDetails">
          <xsl:with-param name ="LayoutName" select ="'page-number'"/>
        </xsl:call-template>
      </p:spPr >
      <p:txBody>
        <a:bodyPr />
        <a:lstStyle />
        <a:p>
          <a:fld type="slidenum">
            <xsl:attribute name ="id">
              <xsl:value-of select ="'{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}'"/>
            </xsl:attribute>
            <a:rPr lang="en-US" smtClean="0" />
            <a:pPr />
            <a:t>
              <xsl:value-of select ="$pageNumber"/>
            </a:t>
          </a:fld>
          <!--<a:endParaRPr lang="en-US" />-->
        </a:p>
      </p:txBody>
    </p:sp >
  </xsl:template>
  <xsl:template name ="tmpNotesfooterDate">
    <xsl:param name ="tmpNotesfooterDateId"/>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="4" name="Date Placeholder 3" />
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="dt" sz="half" idx="10" />
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr >
        <!-- footer layout details style:master-page style:name -->
        <xsl:call-template name ="tmpNotesGetFrameDetails">
          <xsl:with-param name ="LayoutName" select ="'date-time'"/>
        </xsl:call-template>
      </p:spPr >
      <p:txBody>
        <a:bodyPr />
        <a:lstStyle />
        <a:p>
          <xsl:for-each select ="document('content.xml') 
					//presentation:date-time-decl[@presentation:name=$tmpNotesfooterDateId] ">
            <xsl:choose>
              <xsl:when test="@presentation:source='current-date'" >
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
                  <a:rPr lang="en-US" smtClean="0" />
                  <a:t>
                    <xsl:value-of select ="."/>
                  </a:t>
                </a:fld>
                <a:endParaRPr lang="en-US" />
              </xsl:when>
              <xsl:otherwise >
                <a:r>
                  <a:rPr lang="en-US" smtClean="0" />
                  <a:t>
                    <xsl:value-of select ="."/>
                  </a:t>
                </a:r >
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </a:p>
      </p:txBody>
    </p:sp >
  </xsl:template>
  <xsl:template name ="tmpNotesTextAlignment">
    <xsl:param name ="prId"/>
    <a:bodyPr>
      <xsl:for-each select ="draw:text-box">
        <xsl:for-each select ="text:p">
          <xsl:variable name ="ParId">
            <xsl:value-of select ="@text:style-name"/>
          </xsl:variable>
          <xsl:for-each select ="document('content.xml')//style:style[@style:name=$ParId]/style:paragraph-properties">
            <xsl:if test ="@style:writing-mode='tb-rl'">
              <xsl:attribute name ="vert">
                <xsl:value-of select ="'vert'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:for-each>
      <xsl:for-each select ="document('content.xml')//style:style[@style:name=$prId]/style:graphic-properties">
        <xsl:attribute name ="tIns">
          <xsl:if test ="@fo:padding-top and @fo:padding-top!='cm'" >
            <xsl:value-of select ="format-number(substring-before(@fo:padding-top,'cm')  *   360000 ,'#.##')"/>
          </xsl:if>
          <!-- Added by Lohith - or @fo:padding-top = '0cm'-->
          <xsl:if test ="not(@fo:padding-top) or @fo:padding-top = '0cm' or @fo:padding-top='cm'">
            <xsl:value-of select ="0"/>
          </xsl:if >
        </xsl:attribute>
        <xsl:attribute name ="bIns">
          <xsl:if test ="@fo:padding-bottom and @fo:padding-bottom != 'cm'">
            <xsl:value-of select ="format-number(substring-before(@fo:padding-bottom,'cm')  *   360000 ,'#.##')"/>
          </xsl:if>
          <!-- Added by Lohith - or @fo:padding-bottom = '0cm'-->
          <xsl:if test ="not(@fo:padding-bottom) or @fo:padding-bottom = '0cm' or @fo:padding-bottom='cm'">
            <xsl:value-of select ="0"/>
          </xsl:if >
        </xsl:attribute>
        <xsl:attribute name ="lIns">
          <xsl:if test ="@fo:padding-left and @fo:padding-left != 'cm'">
            <xsl:value-of select ="format-number(substring-before(@fo:padding-left,'cm')  *   360000 ,'#.##')"/>
          </xsl:if>
          <!-- Added by Lohith - or @fo:padding-left = '0cm'-->
          <xsl:if test ="not(@fo:padding-left) or @fo:padding-left = '0cm' or @fo:padding-left='cm'">
            <xsl:value-of select ="0"/>
          </xsl:if >
        </xsl:attribute>
        <xsl:attribute name ="rIns">
          <xsl:if test ="@fo:padding-right and @fo:padding-right != 'cm'">
            <xsl:value-of select ="format-number(substring-before(@fo:padding-right,'cm')  *   360000 ,'#.##')"/>
          </xsl:if>
          <!-- Added by Lohith - or @fo:padding-right = '0cm'-->
          <xsl:if test ="not(@fo:padding-right) or @fo:padding-right = '0cm' or @fo:padding-right='cm'">
            <xsl:value-of select ="0"/>
          </xsl:if >
        </xsl:attribute>
        <xsl:variable name ="anchorVal">
          <xsl:choose >
            <xsl:when test ="@draw:textarea-vertical-align ='top'">
              <xsl:value-of select ="'t'"/>
            </xsl:when>
            <xsl:when test ="@draw:textarea-vertical-align ='middle'">
              <xsl:value-of select ="'ctr'"/>
            </xsl:when>
            <xsl:when test ="@draw:textarea-vertical-align ='bottom'">
              <xsl:value-of select ="'b'"/>
            </xsl:when>
            <!--<xsl:otherwise >
							<xsl:value-of select ="'t'"/>
						</xsl:otherwise>-->
          </xsl:choose>
        </xsl:variable>
        <xsl:if test ="$anchorVal != ''">
          <xsl:attribute name ="anchor">
            <xsl:value-of select ="$anchorVal"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name ="anchorCtr">
          <xsl:choose >
            <xsl:when test ="@draw:textarea-horizontal-align ='center'">
              <xsl:value-of select ="1"/>
            </xsl:when>
            <xsl:when test ="@draw:textarea-horizontal-align='justify'">
              <xsl:value-of select ="0"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select ="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name ="wrap">
          <xsl:choose >
            <!--<xsl:when test ="@fo:wrap-option ='no-wrap'">
							<xsl:value-of select ="'none'"/>
						</xsl:when>
						<xsl:when test ="@fo:wrap-option ='wrap'">
							<xsl:value-of select ="'square'"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="'square'"/>
						</xsl:otherwise>-->
            <xsl:when test ="((@draw:auto-grow-height = 'false') and (@draw:auto-grow-width = 'false'))">
              <xsl:value-of select ="'none'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'square'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

      </xsl:for-each>
    </a:bodyPr>
  </xsl:template>
  <xsl:template name ="tmpfontStyles">
    <xsl:param name ="TextStyleID"/>
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl"/>
    <xsl:param name ="masterPageName"/>
    <xsl:message terminate="no">progress:text:p</xsl:message>
  
    <xsl:for-each  select ="document('content.xml')//office:automatic-styles/style:style[@style:name =$TextStyleID ]">
      <xsl:message terminate="no">progress:text:p</xsl:message>
      <!-- Added by lohith :substring-before(style:text-properties/@fo:font-size,'pt')&gt; 0  because sz(font size) shouldnt be zero - 16filesbug-->
      <xsl:if test="style:text-properties/@fo:font-size and substring-before(style:text-properties/@fo:font-size,'pt')&gt; 0 ">
        <xsl:attribute name ="sz">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'pt'"/>
            <xsl:with-param name ="length" select ="style:text-properties/@fo:font-size"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <!--Superscript and SubScript for Text added by Mathi on 31st Jul 2007-->
      <xsl:if test="style:text-properties/@style:text-position">
        <xsl:choose>
          <xsl:when test="(substring-before(style:text-properties/@style:text-position,' ') = 'super')">
            <xsl:attribute name="baseline">
              <xsl:variable name="blsuper">
                <xsl:value-of select="substring-before(substring-after(style:text-properties/@style:text-position,'super '),'%')"/>
              </xsl:variable>
              <xsl:value-of select="($blsuper * 1000)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(substring-before(style:text-properties/@style:text-position,' ') = 'sub')">
            <xsl:attribute name="baseline">
              <xsl:variable name="blsub">
                <xsl:value-of select="substring-before(substring-after(style:text-properties/@style:text-position,'sub '),'%')"/>
              </xsl:variable>
              <xsl:value-of select="($blsub * (-1000))"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:if>
      <!--Font bold attribute -->
      <xsl:if test="style:text-properties/@fo:font-weight[contains(.,'bold')]">
        <xsl:attribute name ="b">
          <xsl:value-of select ="'1'"/>
        </xsl:attribute >
      </xsl:if >
      <!-- Kerning -->
      <xsl:if test ="style:text-properties/@style:letter-kerning = 'true'">
        <xsl:attribute name ="kern">
          <xsl:value-of select="1200"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="style:text-properties/@style:letter-kerning = 'false'">
        <xsl:attribute name ="kern">
          <xsl:value-of select="0"/>
        </xsl:attribute>
      </xsl:if>
      <!-- Font Inclined-->
      <xsl:if test="style:text-properties/@fo:font-style[contains(.,'italic')]">
        <xsl:attribute name ="i">
          <xsl:value-of select ="'1'"/>
        </xsl:attribute >
      </xsl:if >
      
      <!-- Font underline-->
      <xsl:call-template name="tmpNotesUnderLine"/>
      
      <!-- Font Strike through -->
      <xsl:choose >
        <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
          <xsl:attribute name ="strike">
            <xsl:value-of select ="'sngStrike'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
          <xsl:attribute name ="strike">
            <xsl:value-of select ="'dblStrike'"/>
          </xsl:attribute >
        </xsl:when >
        <!-- style:text-line-through-style-->
        <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
          <xsl:attribute name ="strike">
            <xsl:value-of select ="'sngStrike'"/>
          </xsl:attribute >
        </xsl:when>
      </xsl:choose>
      <!--Charector spacing -->
      <xsl:if test ="style:text-properties/@fo:letter-spacing [contains(.,'cm')]">
        <!-- Modfied by lohith - "spc" should be a int value, '#.##'has been replaced by '#'   -->
        <xsl:attribute name ="spc">
          <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
            <!--<xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 3.5 *1000 ,'#')"/>-->
            <xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
          </xsl:if >
          <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
								&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
            <!--<xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') div .035) *100 ,'#')"/>-->
            <xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') * 72 div 2.54) *100 ,'#')"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if >
      <!--Color Node set as standard colors -->
      <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
      <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
      <xsl:if test ="style:text-properties/@fo:color">
        <a:solidFill>
          <a:srgbClr  >
            <xsl:attribute name ="val">
              <!--<xsl:value-of   select ="substring-after(style:text-properties/@fo:color,'#')"/>-->
              <xsl:value-of select ="translate(substring-after(style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
            </xsl:attribute>
          </a:srgbClr >
        </a:solidFill>
      </xsl:if>
      <!-- Text Shadow fix -->
      <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
        <a:effectLst>
          <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
            <a:srgbClr val="000000">
              <a:alpha val="43137" />
            </a:srgbClr>
          </a:outerShdw>
        </a:effectLst>
      </xsl:if>

      <xsl:if test ="style:text-properties/@fo:font-family">
        <a:latin charset="0" >
          <xsl:attribute name ="typeface" >
            <!-- fo:font-family-->
            <xsl:value-of select ="translate(style:text-properties/@fo:font-family, &quot;'&quot;,'')" />
          </xsl:attribute>
        </a:latin >
      </xsl:if>
      <!-- Underline color -->
      <xsl:if test ="style:text-properties/style:text-underline-color">
        <a:uFill>
          <a:solidFill>
            <a:srgbClr>
              <xsl:attribute name ="val">
                <xsl:value-of select ="substring-after(style:text-properties/style:text-underline-color,'#')"/>
              </xsl:attribute>
            </a:srgbClr>
          </a:solidFill>
        </a:uFill>
      </xsl:if>
    </xsl:for-each >
  </xsl:template>
  <xsl:template name="tmpNotesUnderLine">
    <xsl:choose >
      <xsl:when test="style:text-properties/@style:text-underline-type = 'single'">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'sng'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style = 'solid' and
								style:text-properties/@style:text-underline-type[contains(.,'double')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dbl'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style  = 'solid' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'heavy'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style = 'solid' and
							style:text-properties/@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'sng'"/>
        </xsl:attribute >
      </xsl:when>
      <!-- Dotted lean and dotted bold under line -->
      <xsl:when test="style:text-properties/@style:text-underline-style = 'dotted' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dotted'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style = 'dotted' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dottedHeavy'"/>
        </xsl:attribute >
      </xsl:when>
      <!-- Dash lean and dash bold underline -->
      <xsl:when test="style:text-properties/@style:text-underline-style = 'dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dash'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style = 'dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dashHeavy'"/>
        </xsl:attribute >
      </xsl:when>
      <!-- Dash long and dash long bold -->
      <xsl:when test="style:text-properties/@style:text-underline-style = 'long-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dashLong'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style = 'long-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dashLongHeavy'"/>
        </xsl:attribute >
      </xsl:when>

      <!-- dot Dash and dot dash bold -->
      <xsl:when test="style:text-properties/@style:text-underline-style = 'dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dotDash'"/>
          <!-- Modified by lohith for fix 1739785 - dotDashLong to dotDash-->
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style = 'dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dotDashHeavy'"/>
        </xsl:attribute >
      </xsl:when>
      <!-- dot-dot-dash-->
      <xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dotDotDash'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dotDotDashHeavy'"/>
        </xsl:attribute >
      </xsl:when>
      <!-- double Wavy -->
      <xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'wave')] and
								style:text-properties/@style:text-underline-type[contains(.,'double')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'wavyDbl'"/>
        </xsl:attribute >
      </xsl:when>
      <!-- Wavy and Wavy Heavy-->
      <xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'wave')] and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'wavy'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'wave')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'wavyHeavy'"/>
        </xsl:attribute >
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="tmpNotesparaProperties" >
    <!--- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
    <xsl:param name ="paraId" />
    <xsl:param name ="listId"/>
    <xsl:param name ="isBulleted" />
    <xsl:param name ="level"/>
    <xsl:param name ="isNumberingEnabled" />
    <xsl:param name ="framePresentaionStyleId"/>
    <!-- parameter added by vijayeta, dated 13-7-07-->
    <xsl:param name ="masterPageName"/>
    <xsl:param name="slideMaster" />
    <xsl:message terminate="no">progress:text:p</xsl:message>
    <xsl:variable name ="fileName">
      <xsl:if test ="$slideMaster !=''">
        <xsl:value-of select ="$slideMaster"/>
      </xsl:if>
      <xsl:if test ="$slideMaster =''">
        <xsl:value-of select ="'content.xml'"/>
      </xsl:if >
    </xsl:variable >
    <xsl:for-each select ="document($fileName)//style:style[@style:name=$paraId]">
      <xsl:message terminate="no">progress:text:p</xsl:message>
      <a:pPr>
        <!-- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
        <xsl:if test ="not($level='0')">
          <xsl:attribute name ="lvl">
            <xsl:value-of select ="$level"/>
          </xsl:attribute>
        </xsl:if>

        <!--marL="first line indent property"-->
        <xsl:if test ="style:paragraph-properties/@fo:text-indent 
							and substring-before(style:paragraph-properties/@fo:text-indent,'cm') != 0">
          <xsl:attribute name ="indent">
            <!--fo:text-indent-->
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:text-indent"/>
              <xsl:with-param name ="unit" select ="'cm'"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if >
        <xsl:if test ="style:paragraph-properties/@fo:text-align">
          <xsl:attribute name ="algn">
            <!--fo:text-align-->
            <xsl:choose >
              <xsl:when test ="style:paragraph-properties/@fo:text-align='center'">
                <xsl:value-of select ="'ctr'"/>
              </xsl:when>
              <xsl:when test ="style:paragraph-properties/@fo:text-align='end'">
                <xsl:value-of select ="'r'"/>
              </xsl:when>
              <xsl:when test ="style:paragraph-properties/@fo:text-align='justify'">
                <xsl:value-of select ="'just'"/>
              </xsl:when>
              <xsl:otherwise >
                <xsl:value-of select ="'l'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if >
        <!-- Added by Lohith - to set the text alignment using frame properties-->
        <xsl:if test ="not(style:paragraph-properties/@fo:text-align)">
          <xsl:for-each select ="document('content.xml')//style:style[@style:name=$framePresentaionStyleId]">
            <xsl:if test="style:graphic-properties/@draw:textarea-horizontal-align = 'left'">
              <xsl:attribute name ="algn">
                <xsl:value-of select ="'l'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="style:graphic-properties/@draw:textarea-horizontal-align = 'right'">
              <xsl:attribute name ="algn">
                <xsl:value-of select ="'r'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="style:graphic-properties/@draw:textarea-horizontal-align = 'center'">
              <xsl:attribute name ="algn">
                <xsl:value-of select ="'ctr'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:if >

        <xsl:if test ="style:paragraph-properties/@fo:margin-left and 
							   substring-before(style:paragraph-properties/@fo:margin-left,'cm') &gt; 0">
          <xsl:attribute name ="marL">
            <!--fo:margin-left-->
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-left"/>
              <xsl:with-param name ="unit" select ="'cm'"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if >

        <!--Code inserted by Vijayeta For Line Spacing,
            If the line spacing is in terms of Percentage, multiply the value with 1000-->
        <xsl:if test ="style:paragraph-properties/@fo:line-height and 
					substring-before(style:paragraph-properties/@fo:line-height,'%') &gt; 0 and 
					not(substring-before(style:paragraph-properties/@fo:line-height,'%') = 100)">
          <a:lnSpc>
            <a:spcPct>
              <xsl:attribute name ="val">
                <xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:line-height,'%')* 1000,'#.##') "/>
              </xsl:attribute>
            </a:spcPct>
          </a:lnSpc>
        </xsl:if>
        <!--If the line spacing is in terms of Points,multiply the value with 2835-->
        <xsl:if test ="style:paragraph-properties/@style:line-spacing and 
					substring-before(style:paragraph-properties/@style:line-spacing,'cm') &gt; 0">
          <a:lnSpc>
            <a:spcPts>
              <xsl:attribute name ="val">
                <xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-spacing,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:lnSpc>
        </xsl:if>
        <xsl:if test ="style:paragraph-properties/@style:line-height-at-least and 
					substring-before(style:paragraph-properties/@style:line-height-at-least,'cm') &gt; 0 ">
          <a:lnSpc>
            <a:spcPts>
              <xsl:attribute name ="val">
                <xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-height-at-least,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:lnSpc>
        </xsl:if>
        <!--End of Code inserted by Vijayeta For Line Spacing -->
        <!-- Code Added by Vijayeta,for Paragraph Spacing, Before and After
             Multiply the value in cm with 2835
			 date: on 01-06-07-->
        <xsl:if test ="style:paragraph-properties/@fo:margin-top and 
						substring-before(style:paragraph-properties/@fo:margin-top,'cm') &gt; 0 ">
          <a:spcBef>
            <a:spcPts>
              <xsl:attribute name ="val">
                <!--fo:margin-top-->
                <xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-top,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:spcBef >
        </xsl:if>
        <xsl:if test ="style:paragraph-properties/@fo:margin-bottom and 
					    substring-before(style:paragraph-properties/@fo:margin-bottom,'cm') &gt; 0 ">
          <a:spcAft>
            <a:spcPts>
              <xsl:attribute name ="val">
                <!--fo:margin-bottom-->
                <xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-bottom,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:spcAft>
        </xsl:if >
        <!-- Code Added by Vijayeta,for Paragraph Spacing, Before and After-->
        <!--<xsl:if test ="isBulleted='false'">
				<a:buNone/>
        </xsl:if>-->
        <xsl:if test ="$isNumberingEnabled='false'">
          <a:buNone/>
        </xsl:if>
        <xsl:if test ="$isBulleted='true'">
          <xsl:if test ="$isNumberingEnabled='true'">
            <xsl:call-template name ="insertBulletsNumbers" >
              <xsl:with-param name ="listId" select ="$listId"/>
              <xsl:with-param name ="level" select ="$level+1"/>
              <!-- parameter added by vijayeta, dated 13-7-07-->
              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:if>

        <!--Code Inserted by vijayeta,For Bullets and Numbering,Set Level if present-->
        <!-- @@ Code for paragraph tabs -Start-->
        <xsl:call-template name ="paragraphTabstops"/>
        <!-- @@ Code for paragraph tabs -End-->
      </a:pPr>
    </xsl:for-each >
  </xsl:template>
  <!-- Templates added by vijayeta,to get paragraph style name for each of the levels in multilevelled list-->
  <xsl:template name ="getParaStyleName">
    <xsl:param name ="lvl"/>
    <xsl:choose>
      <xsl:when test ="$lvl='0'">
        <xsl:value-of select ="text:list-item/text:p/@text:style-name"/>
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
  <xsl:template name ="getTextNodeForFontStyle">
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl" />
    <xsl:param name ="HyperlinksForBullets" />
    <xsl:param name ="masterPageName"/>
    <xsl:choose>
      <xsl:when test ="./text:p">
        <xsl:for-each select =".">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>

              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <a:r>
                  <a:rPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="tmpfontStyles">
                        <xsl:with-param name ="TextStyleID" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClassName"/>
                        <xsl:with-param name ="lvl" select ="$lvl"/>
                        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="./text:a">
                      <xsl:copy-of select="$HyperlinksForBullets"/>
                    </xsl:if>
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="tmpfontStyles">
                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    </xsl:call-template>
                  </xsl:if>
                  <xsl:if test="./text:a">
                    <xsl:copy-of select="$HyperlinksForBullets"/>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="tmpNotesDefaultFontSize">
    <xsl:param name ="className"/>
    <xsl:param name ="lvl"/>
    <xsl:param name ="masterPageName" />
    <xsl:message terminate="no">progress:text:p</xsl:message>
    <xsl:variable name ="defaultClsName">
      <xsl:call-template name ="getClassName">
        <xsl:with-param name ="clsName" select="$className"/>
        <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
        <xsl:with-param name ="lvl" select ="$lvl"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose >
      <xsl:when test ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-size">
        <xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-size" />
      </xsl:when>
      <xsl:when test ="document('styles.xml')//style:style[@style:name = concat('Standard-',$className)]/style:text-properties/@fo:font-size">
        <xsl:value-of select ="document('styles.xml')//style:style[@style:name = concat('Standard-',$className)]/style:text-properties/@fo:font-size" />
      </xsl:when>
      <xsl:when test="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size">
        <xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@fo:font-size" />
      </xsl:when>
      <xsl:when test="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-asian">
        <xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-asian" />
      </xsl:when>
      <xsl:when test="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-complex">
        <xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-complex" />
      </xsl:when>
      <!-- Added by lohith : sz(font size) shouldnt be zero - 16filesbug-->
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@fo:font-size">
            <xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@fo:font-size" />
          </xsl:when>
          <xsl:when test="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-asian">
            <xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-asian" />
          </xsl:when>
          <xsl:when test="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-complex">
            <xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-complex" />
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpgetCordinates">
    <xsl:for-each select=".">
      <a:xfrm>
        <xsl:variable name ="angle">
          <xsl:value-of select="substring-after(substring-before(substring-before(@draw:transform,'translate'),')'),'(')" />
        </xsl:variable>
        <xsl:variable name ="x2">
          <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
        </xsl:variable>
        <xsl:variable name ="y2">
          <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
        </xsl:variable>
        <xsl:if test="@draw:transform">
          <xsl:attribute name ="rot">
            <xsl:value-of select ="concat('draw-transform:ROT:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>
        </xsl:if>
        <a:off>
          <xsl:if test="@draw:transform">
            <xsl:attribute name ="x">
              <xsl:value-of select ="concat('draw-transform:X:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
            </xsl:attribute>

            <xsl:attribute name ="y">
              <xsl:value-of select ="concat('draw-transform:Y:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="not(@draw:transform)">
            <xsl:attribute name ="x">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="@svg:x"/>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="y">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="@svg:y"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </a:off>
        <a:ext>
          <xsl:attribute name ="cx">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="@svg:width"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name ="cy">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="@svg:height"/>
            </xsl:call-template>
          </xsl:attribute>
        </a:ext>
      </a:xfrm>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="tmpNotesGetFrameDetails">
    <xsl:param name ="LayoutName"/>
    <xsl:for-each select ="document('styles.xml')//style:master-page[@style:name='Default']/draw:frame[@presentation:class=$LayoutName]">
      <a:xfrm>
        <xsl:variable name ="angle">
          <xsl:value-of select="substring-after(substring-before(substring-before(@draw:transform,'translate'),')'),'(')" />
        </xsl:variable>
        <xsl:variable name ="x2">
          <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
        </xsl:variable>
        <xsl:variable name ="y2">
          <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
        </xsl:variable>
        <xsl:if test="@draw:transform">
          <xsl:attribute name ="rot">
            <xsl:value-of select ="concat('draw-transform:ROT:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>
        </xsl:if>
        <a:off>
          <xsl:if test="@draw:transform">
            <xsl:attribute name ="x">
              <xsl:value-of select ="concat('draw-transform:X:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
            </xsl:attribute>

            <xsl:attribute name ="y">
              <xsl:value-of select ="concat('draw-transform:Y:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="not(@draw:transform)">
            <xsl:attribute name ="x">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="@svg:x"/>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="y">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="@svg:y"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </a:off>
        <a:ext>
          <xsl:attribute name ="cx">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="@svg:width"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name ="cy">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="@svg:height"/>
            </xsl:call-template>
          </xsl:attribute>
        </a:ext>
      </a:xfrm>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="NotesRel" match ="/office:document-content/office:body/office:presentation/draw:page">
    <xsl:param name ="slideNo"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

      <xsl:variable name ="slMasterName">
        <xsl:value-of select ="@draw:master-page-name"/>
      </xsl:variable>
      <xsl:variable name ="slMasterLink">
        <xsl:for-each select ="document('styles.xml')//style:master-page/@style:name">
          <xsl:if test ="$slMasterName = .">
            <xsl:value-of select ="position()"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>
      <xsl:variable name ="layoutTemplate">
        <xsl:value-of select ="substring-after(@presentation:presentation-page-layout-name,'T')"/>
      </xsl:variable >
      <Relationship Id="rId2"
							  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide">
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slides/slide',$slideNo,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId1" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster">
        <xsl:attribute name ="Target">
          <xsl:value-of select ="'../notesMasters/notesMaster1.xml'"/>
        </xsl:attribute>
      </Relationship>
    </Relationships>
  </xsl:template>
</xsl:stylesheet>
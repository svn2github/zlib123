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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
  xmlns:odf="urn:odf"
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
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
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  exclude-result-prefixes="odf style text number draw page r presentation fo script xlink svg">
  <xsl:template name="point-measure">
    <xsl:param name="length"/>
    <!-- (string) The length including the unit -->
	
    <xsl:param name="round">true</xsl:param>
    <xsl:variable name="newlength">
      <xsl:choose>
        <xsl:when test="contains($length, 'cm')">
          <xsl:value-of select="number(substring-before($length, 'cm')) * 72 div 2.54"/>
        </xsl:when>
        <xsl:when test="contains($length, 'mm')">
          <xsl:value-of select="number(substring-before($length, 'mm')) * 72 div 25.4"/>
        </xsl:when>
        <xsl:when test="contains($length, 'in')">
          <xsl:value-of select="number(substring-before($length, 'in')) * 72"/>
        </xsl:when>
        <xsl:when test="contains($length, 'pt')">
          <xsl:value-of select="number(substring-before($length, 'pt'))"/>
        </xsl:when>
        <xsl:when test="contains($length, 'pica')">
          <xsl:value-of select="number(substring-before($length, 'pica')) * 12"/>
        </xsl:when>
        <xsl:when test="contains($length, 'dpt')">
          <xsl:value-of select="number(substring-before($length, 'dpt'))"/>
        </xsl:when>
        <xsl:when test="contains($length, 'px')">
          <xsl:value-of select="number(substring-before($length, 'px')) * 72 div 96.19"/>
        </xsl:when>
        <xsl:when test="not($length) or $length='' ">0</xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$length"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$round='true'">
        <xsl:value-of select="round($newlength)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="(round($newlength * 100)) div 100"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="tmpGetTableWidth">
    <xsl:param name="styleName"/>
    <xsl:for-each select="//style:style[@style:name=$styleName]/style:table-column-properties">
      <xsl:if test="@style:column-width">
        <xsl:attribute name ="w">
          <xsl:variable name="convertUnit">
            <xsl:call-template name="getConvertUnit">
              <xsl:with-param name="length" select="@style:column-width"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="$convertUnit"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="@style:column-width"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpGetTableHeight">
    <xsl:param name="styleName"/>
    <xsl:for-each select="//style:style[@style:name=$styleName]/style:table-row-properties">
      <xsl:if test="@style:row-height">
        <xsl:attribute name ="h">
          <xsl:variable name="convertUnit">
            <xsl:call-template name="getConvertUnit">
              <xsl:with-param name="length" select="@style:row-height"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="$convertUnit"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="@style:row-height"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpTable">
    <xsl:param name ="pageNo" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
    <xsl:choose>
      <xsl:when test="document(concat(substring-after(./child::node()[1]/@xlink:href,'./'),'/content.xml'))/child::node() or
                      document(concat(translate(./child::node()[1]/@xlink:href,'/',''),'/content.xml'))/child::node() ">
        <pzip:copy pzip:source="#CER#WordprocessingConverter.dll#OdfConverter.Wordprocessing.resources.OLEplaceholder.png#"
              pzip:target="{concat('ppt/media/','oleObjectImage_',generate-id(),'.png')}"/>
        <p:pic>
          <p:nvPicPr>
            <p:cNvPr id="{$shapeCount + 1 }">
              <xsl:attribute name="name">
                <xsl:value-of select="concat('Picture ',$shapeCount+1)"/>
              </xsl:attribute>
              <xsl:attribute name="descr">
                <xsl:value-of select="concat('oleObjectImage_',generate-id(),'.png')"/>
              </xsl:attribute>
            </p:cNvPr >
            <p:cNvPicPr>
              <a:picLocks noChangeAspect="1">
                <xsl:choose>
                  <xsl:when test="$grpFlag!='true'">
                    <xsl:attribute name="noGrp">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </a:picLocks>
            </p:cNvPicPr>
            <p:nvPr/>
          </p:nvPicPr>
          <p:blipFill>
            <a:blip>
              <xsl:attribute name ="r:embed">
                <xsl:value-of select ="concat('oleObjectImage_',generate-id())"/>
              </xsl:attribute>
            </a:blip >
            <a:stretch>
              <a:fillRect />
            </a:stretch>
          </p:blipFill>
          <p:spPr>
            <xsl:for-each select=".">
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
            </xsl:for-each>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
          </p:spPr>
        </p:pic>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="tmpOLE">
          <xsl:with-param name ="pageNo" select="$pageNo" />
          <xsl:with-param name ="shapeCount" select="$shapeCount" />
          <xsl:with-param name ="grpFlag" select="$grpFlag" />
          <xsl:with-param name ="UniqueId" select="$UniqueId" />
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpTables">
    <xsl:param name ="pageNo" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />

    <xsl:for-each select="./child::node()[1]">
   
        <p:graphicFrame>
          <p:nvGraphicFramePr>
            <p:cNvPr>
              <xsl:attribute name="id">
                <xsl:value-of select="$shapeCount+1"/>
              </xsl:attribute>
              <xsl:attribute name="name">
                <xsl:value-of select="concat('Table ',$shapeCount+1)"/>
              </xsl:attribute>
            </p:cNvPr>
            <p:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="1"/>
            </p:cNvGraphicFramePr>
            <p:nvPr/>
          </p:nvGraphicFramePr>
          <xsl:for-each select="..">
            <xsl:call-template name ="tmpdrawCordinates">
              <xsl:with-param name="grpFlag" select="$grpFlag"/>
              <xsl:with-param name="OLETAB" select="'true'"/>
              
            </xsl:call-template>
          </xsl:for-each>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
              <a:tbl>
                <a:tblPr/>
                <a:tblGrid>
                <xsl:for-each select="table:table-column">
                  <a:gridCol>
                    <xsl:call-template name="tmpGetTableWidth">
                      <xsl:with-param name="styleName">
                        <xsl:value-of select="@table:style-name"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </a:gridCol>
                </xsl:for-each>
                </a:tblGrid>
                <xsl:for-each select="table:table-row">
                  <xsl:variable name="rowNo" select="position()"/>
                  <a:tr h="342900">
                    <xsl:call-template name="tmpGetTableHeight">
                      <xsl:with-param name="styleName">
                        <xsl:value-of select="@table:style-name"/>
                      </xsl:with-param>
                    </xsl:call-template>
                    <xsl:for-each select="table:table-cell | table:covered-table-cell">
                      <xsl:variable name="colNo" select="position()"/>
                      <a:tc>
                        <xsl:if test="name(.)='table:covered-table-cell'">
                          <xsl:attribute name="hMerge">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="@table:number-columns-spanned">
                          <xsl:attribute name="gridSpan">
                            <xsl:value-of select="@table:number-columns-spanned"/>
                          </xsl:attribute>
                        </xsl:if>
                        <a:txBody>
                          <xsl:call-template name ="getParagraphProperties">
                            <xsl:with-param name ="fileName" select ="'content.xml'" />
                            <xsl:with-param name ="gr" select="@table:style-name"/>
                          </xsl:call-template>

                          <xsl:choose>
                            <xsl:when test ="text:p or text:list">
                              <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
                              <xsl:variable name ="masterPageName" select ="./../../../../@draw:master-page-name"/>
                              <xsl:call-template name ="processShapeText">
                                <xsl:with-param name ="fileName" select ="'content.xml'" />
                                <xsl:with-param name="shapeType" select="'table'" />
                                <xsl:with-param name ="shapeCount" select="$shapeCount" />
                                <xsl:with-param name ="grpFlag" select="$grpFlag" />
                                <xsl:with-param name ="UniqueId" select="$UniqueId" />
                                <xsl:with-param name ="gr" select="@table:style-name">
                                </xsl:with-param >
                                <xsl:with-param name ="masterPageName" select ="$masterPageName" />
                              </xsl:call-template>
                            </xsl:when>
                            <xsl:when test ="not(text:p) and not(text:list) and not(text:p or text:list)">
                              <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:call-template name ="processShapeText">
                                <xsl:with-param name ="fileName" select ="'content.xml'" />
                                <xsl:with-param name="shapeType" select="'table'" />
                                <xsl:with-param name ="shapeCount" select="$shapeCount" />
                                <xsl:with-param name ="UniqueId" select="$UniqueId" />
                                <xsl:with-param name ="grpFlag" select="$grpFlag" />
                                <xsl:with-param name ="gr" select="@table:style-name"/>
                              </xsl:call-template>
                            </xsl:otherwise>
                          </xsl:choose>
                        </a:txBody>
                        <a:tcPr>
                        <xsl:call-template name ="getTableGraphicProperties">
                          <xsl:with-param name ="fileName" select ="'content.xml'"/>
                          <xsl:with-param name ="gr">
                            <xsl:choose>
                              <xsl:when test="@table:style-name">
                                <xsl:value-of select="@table:style-name"/>
                              </xsl:when>
                              <xsl:when test="../@table:default-cell-style-name">
                                <xsl:value-of select="../@table:default-cell-style-name"/>
                              </xsl:when>
                            </xsl:choose>
                          </xsl:with-param> 
                          
                          <xsl:with-param name ="useFirstRowStyles" select ="../../@table:use-first-row-styles"/>
                          
                          <xsl:with-param name ="templateName" select ="../../@table:template-name"/>
                          <xsl:with-param name ="rowNo" select ="$rowNo"/>
                          <xsl:with-param name ="colNo" select ="$colNo"/>
                          <xsl:with-param name ="shapeCount" select ="generate-id(.)" />
                          <xsl:with-param name ="grpFlag" select ="$grpFlag" />
                          <xsl:with-param name ="UniqueId" select ="$UniqueId" />
                        </xsl:call-template>
                        </a:tcPr>
                      </a:tc>
                   
                    </xsl:for-each>
                  </a:tr>
                </xsl:for-each>
              </a:tbl>
            </a:graphicData>
          </a:graphic>
        </p:graphicFrame>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template name ="STTextFontSizeInPoints">
    <xsl:param name ="unit"/>
    <xsl:param name ="length"/>
    <xsl:variable name="varSz">
      <xsl:call-template name="convertToPoints">
        <xsl:with-param name="unit" select="$unit"/>
        <xsl:with-param name="length" select="$length"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
			<xsl:when test="$varSz &gt; 400000">
				<xsl:value-of select="400000"/>
      </xsl:when>
			<xsl:when test="$varSz &lt; 100">
				<xsl:value-of select="100"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$varSz"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template >
  
	<xsl:template name ="convertToPoints">
		<xsl:param name ="unit"/>
		<xsl:param name ="length"/>
		<xsl:message terminate="no">progress:text:p</xsl:message>
		<xsl:variable name="lengthVal">
			<xsl:choose>
				<xsl:when test="contains($length,'cm')">
					<xsl:value-of select="substring-before($length,'cm')"/>
				</xsl:when>
				<xsl:when test="contains($length,'pt')">
					<xsl:value-of select="substring-before($length,'pt')"/>
				</xsl:when>
				<xsl:when test="contains($length,'in')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'in') * 2.54 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'in')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!--mm to cm -->
				<xsl:when test="contains($length,'mm')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'mm') div 10 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'mm')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- m to cm -->
				<xsl:when test="contains($length,'m')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'m') * 100 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'m')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- km to cm -->
				<xsl:when test="contains($length,'km')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'km') * 100000  "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'km')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- mi to cm -->
				<xsl:when test="contains($length,'mi')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'mi') * 160934.4"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'mi')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- ft to cm -->
				<xsl:when test="contains($length,'ft')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'ft') * 30.48 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'ft')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- em to cm -->
				<xsl:when test="contains($length,'em')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round(substring-before($length,'em') * .4233) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'em')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- px to cm -->
				<xsl:when test="contains($length,'px')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round(substring-before($length,'px') div 35.43307) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'px')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- pc to cm -->
				<xsl:when test="contains($length,'pc')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round(substring-before($length,'pc') div 2.362) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'pc')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- ex to cm 1 ex to 6 px-->
				<xsl:when test="contains($length,'ex')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round((substring-before($length,'ex') div 35.43307)* 6) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'ex')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$length"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test ="contains($lengthVal,'NaN')">
				<xsl:value-of select ="0"/>
			</xsl:when>
			<xsl:when test="$lengthVal='0' or $lengthVal='' or ( ($lengthVal &lt; 0) and ($unit != 'cm')) ">
				<xsl:value-of select="0"/>
			</xsl:when>
			<xsl:when test="$unit = 'cm'">
				<xsl:value-of select="concat(format-number($lengthVal * 360000,'#'),'')"/>
			</xsl:when>
			<xsl:when test="$unit = 'mm'">
				<xsl:value-of select="concat(format-number($lengthVal * 25.4 div 72,'#.###'),'mm')"/>
			</xsl:when>
			<xsl:when test="$unit = 'in'">
				<xsl:value-of select="concat(format-number($lengthVal div 72,'#.###'),'in')"/>
			</xsl:when>
			<xsl:when test="$unit = 'pt'">
				<xsl:value-of select="concat(round($lengthVal * 100) ,'')"/>
				<!--Added by lohith - format-number($lengthVal,'#') to make sure that pt will be a int not a real value-->
			</xsl:when>
			<xsl:when test="$unit = 'pica'">
				<xsl:value-of select="concat(format-number($lengthVal div 12,'#.###'),'pica')"/>
			</xsl:when>
			<xsl:when test="$unit = 'dpt'">
				<xsl:value-of select="concat($lengthVal,'dpt')"/>
			</xsl:when>
			<xsl:when test="$unit = 'px'">
				<xsl:value-of select="concat(format-number($lengthVal * 96.19 div 72,'#.###'),'px')"/>
			</xsl:when>
			<xsl:when test="not($lengthVal)">
				<xsl:value-of select="concat(0,'cm')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$lengthVal"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
  <xsl:template name ="convertUnitsToCm">
    <xsl:param name ="unit"/>
    <xsl:param name ="length"/>
    <xsl:choose>
      <xsl:when test="contains($length,'cm')">
        <xsl:value-of select="substring-before($length,'cm')"/>
      </xsl:when>
      <xsl:when test="contains($length,'pt')">
        <xsl:value-of select="substring-before($length,'pt') * 0.035277778"/>
      </xsl:when>
      <xsl:when test="contains($length,'in')">
        <xsl:value-of select="substring-before($length,'in') * 2.54 "/>
      </xsl:when>
      <!--mm to cm -->
      <xsl:when test="contains($length,'mm')">
        <xsl:value-of select="substring-before($length,'mm') div 10 "/>
      </xsl:when>
          <!-- km to cm -->
      <xsl:when test="contains($length,'km')">
        <xsl:value-of select="substring-before($length,'km') * 100000  "/>
      </xsl:when>
      <!-- mi to cm -->
      <xsl:when test="contains($length,'mi')">
        <xsl:value-of select="substring-before($length,'mi') * 160934.4"/>
      </xsl:when>

      <!-- ft to cm -->
      <xsl:when test="contains($length,'ft')">
        <xsl:value-of select="substring-before($length,'ft') * 30.48 "/>
      </xsl:when>
      <!-- em to cm -->
      <xsl:when test="contains($length,'em')">
        <xsl:value-of select="round(substring-before($length,'em') * .4233) "/>
      </xsl:when>
      <!-- px to cm -->
      <xsl:when test="contains($length,'px')">
        <xsl:value-of select="round(substring-before($length,'px') div 35.43307) "/>
      </xsl:when>
      <!-- pc to cm -->
      <xsl:when test="contains($length,'pc')">
        <xsl:value-of select="round(substring-before($length,'pc') div 2.362) "/>
      </xsl:when>
      <!-- ex to cm 1 ex to 6 px-->
      <xsl:when test="contains($length,'ex')">
        <xsl:value-of select="round((substring-before($length,'ex') div 35.43307)* 6) "/>
      </xsl:when>
      <!-- m to cm -->
      <xsl:when test="contains($length,'m')">
        <xsl:value-of select="substring-before($length,'m') * 100 "/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template >
  <xsl:template name ="tmpconvertToPoints">
    <xsl:param name ="unit"/>
    <xsl:param name ="length"/>
    <xsl:message terminate="no">progress:text:p</xsl:message>
    <xsl:variable name="lengthVal">
      <xsl:choose>
        <xsl:when test="contains($length,'cm')">
          <xsl:value-of select="substring-before($length,'cm')"/>
        </xsl:when>
        <xsl:when test="contains($length,'pt')">
          <xsl:value-of select="substring-before($length,'pt')"/>
        </xsl:when>
        <xsl:when test="contains($length,'in')">
          <xsl:value-of select="substring-before($length,'in') * 2.54 "/>
        </xsl:when>
        <!--mm to cm -->
        <xsl:when test="contains($length,'mm')">
          <xsl:value-of select="substring-before($length,'mm') div 10 "/>
        </xsl:when>
        <!-- m to cm -->
        <xsl:when test="contains($length,'m')">
          <xsl:value-of select="substring-before($length,'m') * 100 "/>
        </xsl:when>
        <!-- km to cm -->
        <xsl:when test="contains($length,'km')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'km') * 100000  "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'km')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- mi to cm -->
        <xsl:when test="contains($length,'mi')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'mi') * 160934.4"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'mi')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- ft to cm -->
        <xsl:when test="contains($length,'ft')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'ft') * 30.48 "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'ft')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- em to cm -->
        <xsl:when test="contains($length,'em')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round(substring-before($length,'em') * .4233) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'em')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- px to cm -->
        <xsl:when test="contains($length,'px')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round(substring-before($length,'px') div 35.43307) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'px')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- pc to cm -->
        <xsl:when test="contains($length,'pc')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round(substring-before($length,'pc') div 2.362) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'pc')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- ex to cm 1 ex to 6 px-->
        <xsl:when test="contains($length,'ex')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round((substring-before($length,'ex') div 35.43307)* 6) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'ex')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$length"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test ="contains($lengthVal,'NaN')">
        <xsl:value-of select ="0"/>
      </xsl:when>
      <xsl:when test="$lengthVal='0' or $lengthVal='' or ( ($lengthVal &lt; 0) and ($unit != 'cm')) ">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($lengthVal * 360000,'#'),'')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($lengthVal * 25.4 div 72,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($lengthVal div 72,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat(format-number($lengthVal,'#') * 100 ,'')"/>
        <!--Added by lohith - format-number($lengthVal,'#') to make sure that pt will be a int not a real value-->
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($lengthVal div 12,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat($lengthVal,'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($lengthVal * 96.19 div 72,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($lengthVal)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$lengthVal"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template >
	<xsl:template name ="fontStyles">
		<xsl:param name ="Tid"/>
		<xsl:param name ="prClassName"/>
    <xsl:param name ="flagPresentationClass"/>
		<xsl:param name ="lvl"/>
    <xsl:param name="parentStyleName"/>
		<!-- Parameter Added by Vijayeta, on 13-7-07-->
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

		<xsl:for-each  select ="document($fileName)//office:automatic-styles/style:style[@style:name =$Tid ]">
      <xsl:if test="position()=1">
			<xsl:message terminate="no">progress:text:p</xsl:message>

        <xsl:choose>
          <xsl:when test="style:text-properties/@style:language-asian and style:text-properties/@style:country-asian">
            <xsl:attribute name ="lang">
              <xsl:value-of select="concat(style:text-properties/@style:language-asian,'-',style:text-properties/@style:country-asian)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="lang">
              <xsl:value-of select="'en-US'"/>
            </xsl:attribute>
        </xsl:otherwise>
        </xsl:choose>
			<!-- Added by lohith :substring-before(style:text-properties/@fo:font-size,'pt')&gt; 0  because sz(font size) shouldnt be zero - 16filesbug-->
<!--Office 2007 Sp2-->
        <xsl:variable name="fontSize">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:if test="$fontSize &gt; 0 ">
				<xsl:attribute name ="sz">
					<xsl:call-template name ="STTextFontSizeInPoints">
						<xsl:with-param name ="unit" select ="'pt'"/>
						<xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<!--Superscript and SubScript for Text added by Mathi on 31st Jul 2007-->
			<xsl:if test="style:text-properties/@style:text-position">
        <xsl:call-template name="tmpSuperSubScriptForward"/>
			</xsl:if>

			<xsl:if test ="not(style:text-properties/@fo:font-size) and ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param> 
          <xsl:with-param name="attrName" select="'Fontsize'"/>
					</xsl:call-template >
			      </xsl:if>
			<!--Font bold attribute -->
        <xsl:choose>
          <xsl:when test="style:text-properties/@fo:font-weight[contains(.,'bold')]">
				<xsl:attribute name ="b">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
          </xsl:when >
          <xsl:when test="style:text-properties/@fo:font-weight[contains(.,'normal')]">
            <xsl:attribute name ="b">
              <xsl:value-of select ="'0'"/>
            </xsl:attribute >
          </xsl:when >
          <xsl:when test ="not(style:text-properties/@fo:font-weight[contains(.,'bold')]) and ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="attrName" select="'Bold'"/>
        </xsl:call-template>
          </xsl:when>
        </xsl:choose>
			<!-- Kerning - Added by lohith.ar -->
			<!-- Start -->
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
      <xsl:if test ="not(style:text-properties/@style:letter-kerning) and   ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="attrName" select="'kerning'"/>
        </xsl:call-template>
      </xsl:if>
			
			<!-- End -->
			<!-- Font Inclined-->
      <xsl:choose>
        <xsl:when test="style:text-properties/@fo:font-style='italic'">
				<xsl:attribute name ="i">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
        </xsl:when >
        <xsl:when test="style:text-properties/@fo:font-style='normal'">
          <xsl:attribute name ="i">
            <xsl:value-of select ="'0'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test="not(style:text-properties/@fo:font-style='normal') and not(style:text-properties/@fo:font-style='italic')
                           and not(style:text-properties/@fo:font-style) and ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="attrName" select="'italic'"/>
        </xsl:call-template>
        </xsl:when >
      </xsl:choose>
			<!-- Font underline-->
    
        <xsl:call-template name="tmpUnderLineStyle">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="flagPresentationClass" select="$flagPresentationClass"/>
          <xsl:with-param name="prClassName" select="$prClassName"/>
        </xsl:call-template>
      

      <!-- Font Strike through Start-->
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
        <xsl:otherwise>
          <xsl:if test="$flagPresentationClass='No' or $prClassName='subtitle'">
            <xsl:call-template name="tmpgetDefualtTextProp">
              <xsl:with-param name="parentStyleName">
                <xsl:choose>
                  <xsl:when test="$prClassName='subtitle'">
                    <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$parentStyleName"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
              <xsl:with-param name="attrName" select="'strike'"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:otherwise>
			</xsl:choose>
			<!-- Font Strike through end-->
			<!--Charector spacing -->
			<!-- Modfied by lohith - @fo:letter-spacing will have a text value 'normal' when no change is required -->
      <xsl:call-template name="tmpCharacterSpacing"/>
			<!--Color Node set as standard colors -->
			<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
			<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
			<xsl:if test ="style:text-properties/@fo:color">
				<a:solidFill>
            <xsl:variable name="varSrgbVal">
              <xsl:value-of select ="translate(substring-after(style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
            </xsl:variable>
            <xsl:if test="$varSrgbVal != ''">
					<a:srgbClr  >
						<xsl:attribute name ="val">
                  <xsl:value-of select ="$varSrgbVal"/>
						</xsl:attribute>
					</a:srgbClr >
            </xsl:if>            
				</a:solidFill>
			</xsl:if>
      <xsl:if test ="not(style:text-properties/@fo:color) and ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="attrName" select="'Fontcolor'"/>
        </xsl:call-template>
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
      <xsl:if test ="not(style:text-properties/@fo:text-shadow) and ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="attrName" select="'Textshadow'"/>
        </xsl:call-template>
      </xsl:if>

      <!-- Underline color -->
      <xsl:if test ="style:text-properties/@style:text-underline-color">
              <xsl:choose>
                <xsl:when test="style:text-properties/@style:text-underline-color='font-color'">
					<xsl:choose>
						<xsl:when test="style:text-properties/@fo:color">
        <a:uFill>
          <a:solidFill>
											<xsl:variable name="varSrgbVal">
												<xsl:value-of select ="translate(substring-after(style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
											</xsl:variable>
											<xsl:if test="$varSrgbVal != ''">
            <a:srgbClr>             
                  <xsl:attribute name ="val">                    
														<xsl:value-of select ="$varSrgbVal"/>
										</xsl:attribute>
									</a:srgbClr>
											</xsl:if>
								</a:solidFill>
							</a:uFill>
						</xsl:when>
						<xsl:when test="not(style:text-properties/@fo:color) and document('styles.xml')//style:style[@style:name = $parentStyleName]/style:text-properties/@fo:color">
                      <xsl:for-each select="document('styles.xml')//style:style[@style:name = $parentStyleName]/style:text-properties">
								<xsl:if test="position()=1">
									<a:uFill>
										<a:solidFill>
													<xsl:variable name="varSrgbVal">
														<xsl:value-of select ="translate(substring-after(@fo:color,'#'),$lcletters,$ucletters)"/>
													</xsl:variable>
													<xsl:if test="$varSrgbVal != ''">
											<a:srgbClr>
												<xsl:attribute name ="val">
																<xsl:value-of select ="$varSrgbVal"/>
                  </xsl:attribute>
											</a:srgbClr>
													</xsl:if>
										</a:solidFill>
									</a:uFill>
								</xsl:if>
							</xsl:for-each>
						</xsl:when>
					</xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>             
                </xsl:otherwise>
              </xsl:choose>          
      </xsl:if>
      <xsl:if test ="not(style:text-properties/@style:text-underline-color) and ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="attrName" select="'Underlinecolor'"/>
        </xsl:call-template>
      </xsl:if>
		
			<xsl:if test ="style:text-properties/@fo:font-family">
				<a:latin charset="0" >
					<xsl:attribute name ="typeface" >
						<!-- fo:font-family-->
						<xsl:value-of select ="translate(style:text-properties/@fo:font-family, &quot;'&quot;,'')" />
					</xsl:attribute>
				</a:latin >
			</xsl:if>
			<!--Commented by Vipul As no need to assign defualt font as Arial-->
			<!--Start-->
      <xsl:if test ="not(style:text-properties/@fo:font-family) and ($flagPresentationClass='No' or $prClassName='subtitle')">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName">
            <xsl:choose>
              <xsl:when test="$prClassName='subtitle'">
                <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$parentStyleName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="attrName" select="'Fontname'"/>
							</xsl:call-template >
			</xsl:if>
			<!--End-->
      </xsl:if>
			</xsl:for-each >
	</xsl:template>
  <xsl:template name="tmpCharacterSpacing">
    <xsl:if test ="style:text-properties/@fo:letter-spacing">
      <xsl:variable name ="spc">
        <xsl:variable name="spcTemp">
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length"  select ="style:text-properties/@fo:letter-spacing"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test ="$spcTemp &lt; 0 ">
          <xsl:value-of select ="format-number($spcTemp * 7200 div 2.54 ,'#')"/>
        </xsl:if >
        <xsl:if test ="$spcTemp &gt; 0 or $spcTemp = 0 ">
          <xsl:value-of select ="format-number(($spcTemp * 72 div 2.54) *100 ,'#')"/>
        </xsl:if>
      </xsl:variable>
      <xsl:if test ="$spc!=''">
        <xsl:attribute name ="spc">
          <xsl:value-of select ="$spc"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:if >
  </xsl:template>
	<xsl:template name ="paraProperties" >
		<!--- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
		<xsl:param name ="paraId" />
    <xsl:param name ="BuImgRel" />
		<xsl:param name ="listId"/>
		<xsl:param name ="isBulleted" />
		<xsl:param name ="level"/>
		<xsl:param name ="isNumberingEnabled" />
		<xsl:param name ="framePresentaionStyleId"/>
    <xsl:param name ="prClsName"/>
		<!-- parameter added by vijayeta, dated 13-7-07-->
		<xsl:param name ="masterPageName"/>
		<xsl:param name="slideMaster" />
    <xsl:param name="pos" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="FrameCount"/>
    <xsl:param name ="flagPresentationClass"/>
    <xsl:param name="parentStyleName"/>
    <xsl:param name ="grpFlag"/>
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
      <xsl:if test="position()=1">
			<xsl:message terminate="no">progress:text:p</xsl:message>
			<a:pPr>
				<!-- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
				<xsl:if test ="not($level='0')">
					<xsl:attribute name ="lvl">
						<xsl:value-of select ="$level"/>
					</xsl:attribute>
				</xsl:if>
				<!-- Added by vijayeta for teh fix 1739081-->
				<xsl:if test ="$isNumberingEnabled='true'">
					<xsl:variable name ="marL">
						<xsl:call-template name="MarginTemplateForSlide">
							<xsl:with-param name="level" select="$level"/>
							<xsl:with-param name ="listId" select ="$listId"/>
							<xsl:with-param name ="fileName" select ="$fileName"/>
						</xsl:call-template >
					</xsl:variable>
					<xsl:if test ="$marL!=''">
						<xsl:attribute name ="marL">
							<xsl:value-of select ="$marL"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:variable name ="indetValue">
						<xsl:call-template name="IndentTemplateForSlide">
							<xsl:with-param name="level" select="$level"/>
							<xsl:with-param name ="listId" select ="$listId"/>
							<xsl:with-param name ="fileName" select ="$fileName"/>
						</xsl:call-template >
					</xsl:variable>
					<xsl:if test ="$indetValue!=''">
						<xsl:attribute name="indent">
                <xsl:value-of select="concat('-',$indetValue)"/>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				<!-- Added by vijayeta for teh fix 1739081-->
				<!--marL="first line indent property"-->
        <!-- Condition which checks if text indent is greater than 0 is removed, as the input has an indent of value 0
             bug number 1779336,by vijayeta, date:23rd aug '07-->
        <xsl:if test ="style:paragraph-properties/@fo:text-indent">
          <!--fo:text-indent-->
          <xsl:variable name ="varIndent">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name="length" select="style:paragraph-properties/@fo:text-indent"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:if test ="$varIndent!=''">
            <xsl:attribute name ="indent">
              <xsl:value-of select ="$varIndent"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if >
        <xsl:if test ="not(style:paragraph-properties/@fo:text-indent) and  ($flagPresentationClass='No' or $prClsName='subtitle')">
          <xsl:call-template name="tmpgetDefualtParagraphProp">
            <xsl:with-param name="parentStyleName">
              <xsl:choose>
                <xsl:when test="$prClsName='subtitle'">
                  <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$parentStyleName"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="attrName" select="'Textindent'"/>
          </xsl:call-template>
        </xsl:if>
        <!-- End,bug number 1779336 vijayeta, date:23rd aug '07-->
				<xsl:if test ="style:paragraph-properties/@fo:text-align">
					<xsl:attribute name ="algn">
						<!--fo:text-align-->
						<xsl:choose >
							<xsl:when test ="style:paragraph-properties/@fo:text-align='center'">
								<xsl:value-of select ="'ctr'"/>
							</xsl:when>
							<xsl:when test ="style:paragraph-properties/@fo:text-align='end' or style:paragraph-properties/@fo:text-align='right'">
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
        <xsl:if test ="not(style:paragraph-properties/@fo:text-align) and  ($flagPresentationClass='No' or $prClsName='subtitle')">
          <xsl:call-template name="tmpgetDefualtParagraphProp">
            <xsl:with-param name="parentStyleName">
              <xsl:choose>
                <xsl:when test="$prClsName='subtitle'">
          <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
              </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$parentStyleName"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="attrName" select="'Textalign'"/>
          </xsl:call-template>
            </xsl:if>
                <!-- Condition which checks if text indent is greater than 0 is removed, as the input has a marginn left of value 0
             bug number 1779336 by vijayeta, date:23rd aug '07-->
        <xsl:call-template name ="tmpMarLeft"/>
        <xsl:if test ="style:paragraph-properties/@fo:margin-right">
          <!-- warn if indent after text-->
          <xsl:message terminate="no">translation.odf2oox.paragraphIndentTypeAfterText</xsl:message>
        </xsl:if>
        <xsl:if test ="not(style:paragraph-properties/@fo:margin-left) and  ($flagPresentationClass='No' or $prClsName='subtitle')">
          <xsl:call-template name="tmpgetDefualtParagraphProp">
            <xsl:with-param name="parentStyleName">
              <xsl:choose>
                <xsl:when test="$prClsName='subtitle'">
                  <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$parentStyleName"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="attrName" select="'MarginLeft'"/>
          </xsl:call-template>
        </xsl:if>
        <!-- End,bug number 1779336 vijayeta, date:23rd aug '07-->
				<!--Code inserted by Vijayeta For Line Spacing,
            If the line spacing is in terms of Percentage, multiply the value with 1000-->
        <xsl:variable name="lineSpc">
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:letter-spacing"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="lineSpcAtleast">
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length"  select ="style:paragraph-properties/@style:line-height-at-least"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>         
          <xsl:when test ="style:paragraph-properties/@fo:line-height and 
					substring-before(style:paragraph-properties/@fo:line-height,'%') &gt; 0 and 
					not(substring-before(style:paragraph-properties/@fo:line-height,'%') = 100)">
					<a:lnSpc>
						<a:spcPct>
							<xsl:attribute name ="val">
								<xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:line-height,'%')* 1000,'#.##') "/>
							</xsl:attribute>
						</a:spcPct>
					</a:lnSpc>
          </xsl:when>
				<!--If the line spacing is in terms of Points,multiply the value with 2835-->
        
          <xsl:when test ="$lineSpc > 0">
					<a:lnSpc>
						<a:spcPts>
              <xsl:attribute name ="val">
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length">
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="$lineSpc"/>
                    </xsl:call-template>
                  </xsl:with-param>
                </xsl:call-template>
                  </xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
          </xsl:when>
          <xsl:when test ="$lineSpcAtleast > 0 ">
					<a:lnSpc>
						<a:spcPts>
              <xsl:attribute name ="val">
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length">
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="$lineSpcAtleast"/>
                    </xsl:call-template>
                  </xsl:with-param>
                </xsl:call-template>
                 </xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="$flagPresentationClass='No' or $prClsName='subtitle'">
              <xsl:call-template name="tmpgetDefualtParagraphProp">
                <xsl:with-param name="parentStyleName">
                  <xsl:choose>
                    <xsl:when test="$prClsName='subtitle'">
                      <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$parentStyleName"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="attrName" select="'Linespacing'"/>
              </xsl:call-template>
				</xsl:if>
          </xsl:otherwise>
        </xsl:choose>

				<!--End of Code inserted by Vijayeta For Line Spacing -->
				<!-- Code Added by Vijayeta,for Paragraph Spacing, Before and After
             Multiply the value in cm with 2835
			 date: on 01-06-07-->
        <xsl:call-template name ="tmpMarTop"/>
        <xsl:if test ="not(style:paragraph-properties/@fo:margin-top) and  ($flagPresentationClass='No' or $prClsName='subtitle')">
          <xsl:call-template name="tmpgetDefualtParagraphProp">
            <xsl:with-param name="parentStyleName">
              <xsl:choose>
                <xsl:when test="$prClsName='subtitle'">
                  <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$parentStyleName"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="attrName" select="'MarginTop'"/>
          </xsl:call-template>
        </xsl:if>
        <xsl:call-template name ="tmpMarBottom"/>
        <xsl:if test ="not(style:paragraph-properties/@fo:margin-bottom) and  ($flagPresentationClass='No' or $prClsName='subtitle')">
          <xsl:call-template name="tmpgetDefualtParagraphProp">
            <xsl:with-param name="parentStyleName">
              <xsl:choose>
                <xsl:when test="$prClsName='subtitle'">
                  <xsl:value-of select="concat($masterPageName,'-subtitle')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$parentStyleName"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="attrName" select="'MarginBottom'"/>
          </xsl:call-template>
        </xsl:if>
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
              <xsl:with-param name ="pos" select ="$pos"/>
              <xsl:with-param name ="grpFlag" select ="$grpFlag"/>
                <xsl:with-param name ="BuImgRel" select ="$BuImgRel"/>
              <xsl:with-param name ="shapeCount" select ="$shapeCount" />
              <xsl:with-param name ="FrameCount" select ="$FrameCount"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:if>
				<!--Code Inserted by vijayeta,For Bullets and Numbering,Set Level if present-->
				<!-- @@ Code for paragraph tabs -Start-->
				<xsl:call-template name ="paragraphTabstops"/>
				<!-- @@ Code for paragraph tabs -End-->
			</a:pPr>
      </xsl:if>
		</xsl:for-each >
	</xsl:template>
  <xsl:template name="tmpMarLeft">
    <xsl:if test ="style:paragraph-properties/@fo:margin-left">
      <!--fo:margin-left-->
      <xsl:variable name ="varMarginLeft">
        <xsl:call-template name ="convertToPoints">
          <xsl:with-param name ="unit" select ="'cm'"/>
          <xsl:with-param name ="length">
            <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-left"/>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test ="$varMarginLeft!=''">
        <xsl:attribute name ="marL">
          <xsl:value-of select ="$varMarginLeft"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:if >
  </xsl:template>
  <xsl:template name="tmpMarTop">
    <xsl:if test ="style:paragraph-properties/@fo:margin-top">
      <xsl:if test ="style:paragraph-properties/@fo:margin-top">
        <a:spcBef>
          <a:spcPts>
            <xsl:attribute name ="val">
              <!--fo:margin-top-->
              <xsl:call-template name ="convertToPointsLineSpacing">
                <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-top"/>
                <xsl:with-param name ="unit" select ="'cm'"/>
              </xsl:call-template>
            </xsl:attribute>
          </a:spcPts>
        </a:spcBef >
      </xsl:if>
    </xsl:if>
     </xsl:template>
  <xsl:template name="tmpMarBottom">
    <xsl:if test ="style:paragraph-properties/@fo:margin-bottom">
      <a:spcAft>
        <a:spcPts>
          <xsl:attribute name ="val">
            <!--fo:margin-bottom-->
            <xsl:call-template name ="convertToPointsLineSpacing">
              <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-bottom"/>
              <xsl:with-param name ="unit" select ="'cm'"/>
            </xsl:call-template>
          </xsl:attribute>
        </a:spcPts>
      </a:spcAft>
    </xsl:if >
  </xsl:template>
 
  <xsl:template name="tmpLineSpacing">
    <xsl:variable name="Unit">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="style:paragraph-properties/@style:line-spacing"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="Unit1">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="style:paragraph-properties/@style:line-height-at-least"/>
      </xsl:call-template>
    </xsl:variable>
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
					substring-before(style:paragraph-properties/@style:line-spacing,$Unit) &gt; 0">
      <xsl:variable name="lineSpacing">
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name ="length" select ="style:paragraph-properties/@style:line-spacing"/>
        </xsl:call-template>
      </xsl:variable>
      <a:lnSpc>
        <a:spcPts>
          <xsl:attribute name ="val">
            <xsl:value-of select ="round($lineSpacing* 2835) "/>
          </xsl:attribute>
        </a:spcPts>
      </a:lnSpc>
    </xsl:if>
    <xsl:if test ="style:paragraph-properties/@style:line-height-at-least and 
					substring-before(style:paragraph-properties/@style:line-height-at-least,$Unit1) &gt; 0 ">
      <xsl:variable name="lineSpacing">
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name ="length" select ="style:paragraph-properties/@style:line-height-at-least"/>
        </xsl:call-template>
      </xsl:variable>
      <a:lnSpc>
        <a:spcPts>
          <xsl:attribute name ="val">
            <xsl:value-of select ="round($lineSpacing* 2835) "/>
          </xsl:attribute>
        </a:spcPts>
      </a:lnSpc>
    </xsl:if>
  </xsl:template>
	<xsl:template name ="fillColor">
		<xsl:param name ="prId"/>
    <xsl:param name ="var_pos"/>
    <xsl:param name="flagFile"/>
    <xsl:param name="opacity"/>

    <xsl:variable name="filename">
      <xsl:choose>
        <xsl:when test="$flagFile=''">
          <xsl:value-of select="'content.xml'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'styles.xml'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
		<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
		<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
		<xsl:for-each select ="document($filename)//style:style[@style:name=$prId] ">
      <xsl:if test="position()=1">
			<!--test="not(style:graphic-properties/@draw:fill = 'none' - Added by lohith.ar for invalid fill color for textboxes - Fill type should be given priority on fill color-->
        <xsl:variable name="parentStyle" select="@style:parent-style-name"/>
      <xsl:choose>
        <xsl:when test ="style:graphic-properties/@draw:fill='solid'">
          <a:solidFill>
              <xsl:variable name="varSrgbVal">
                <xsl:value-of select ="translate(substring-after(style:graphic-properties/@draw:fill-color,'#'),$lcletters,$ucletters)"/>
              </xsl:variable>
              <xsl:if test="$varSrgbVal != ''">
            <a:srgbClr  >
              <xsl:attribute name ="val">
                    <xsl:value-of select ="$varSrgbVal"/>
              </xsl:attribute>
                <xsl:choose>
                  <xsl:when test="$opacity!=''">
                    <xsl:call-template name="tmpshapeTransperancy">
                      <xsl:with-param name="tranparency" select="$opacity"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:otherwise>
              <xsl:if test ="style:graphic-properties/@draw:opacity">
                <xsl:variable name="tranparency" select="substring-before(style:graphic-properties/@draw:opacity,'%')"/>
                <xsl:call-template name="tmpshapeTransperancy">
                  <xsl:with-param name="tranparency" select="$tranparency"/>
                </xsl:call-template>
              </xsl:if>
                  </xsl:otherwise>
                </xsl:choose>
            </a:srgbClr >
              </xsl:if>              
          </a:solidFill>
        </xsl:when>
        <xsl:when test ="style:graphic-properties/@draw:fill='none'">
          <a:noFill/>
          </xsl:when>
        <xsl:when test ="style:graphic-properties/@draw:fill='gradient'">
          <xsl:call-template name="tmpGradientFill">
            <xsl:with-param name="gradStyleName" select="style:graphic-properties/@draw:fill-gradient-name"/>
            <xsl:with-param  name="opacity" select="substring-before(style:graphic-properties/@draw:opacity,'%')"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test ="style:graphic-properties/@draw:fill='bitmap'">
          <xsl:call-template name="tmpBitmapFill">
            <xsl:with-param name="FileName" select="concat('bitmap',$var_pos)" />
            <xsl:with-param name="var_imageName" select="style:graphic-properties/@draw:fill-image-name" />
            <xsl:with-param  name="opacity" select="substring-before(style:graphic-properties/@draw:opacity,'%')"/>
            <xsl:with-param  name="stretch" select="style:graphic-properties/@style:repeat"/>
        </xsl:call-template>
        </xsl:when>
          <xsl:when test ="not(style:graphic-properties/@draw:fill) and $parentStyle!=''">
            <xsl:call-template name="fillColor">
              <xsl:with-param name ="prId" select ="$parentStyle" />
              <xsl:with-param name ="var_pos" select ="$var_pos" />
              <xsl:with-param name ="flagFile" select ="'styles.xml'" />
              <xsl:with-param name ="opacity" >
                <xsl:if test ="style:graphic-properties/@draw:opacity">
                  <xsl:value-of select="substring-before(style:graphic-properties/@draw:opacity,'%')"/>
                </xsl:if>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
      </xsl:choose>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template name="tmpBitmapFill">
    <xsl:param name="var_imageName"/>
    <xsl:param name="FileName"/>
    <xsl:param name="opacity"/>
    <xsl:param name="stretch"/>
   
    <xsl:if test="document('styles.xml')/office:document-styles/office:styles/draw:fill-image[@draw:name=$var_imageName]">
      <a:blipFill dpi="0" rotWithShape="1">
        <xsl:call-template name="tmpInsertBackImage">
          <xsl:with-param name="FileName" select="$FileName" />
          <xsl:with-param name="imageName" select="$var_imageName" />
          <xsl:with-param name="fillType" select="'shape'" />
          <xsl:with-param name="opacity" select="$opacity" />
        </xsl:call-template>
        
        <xsl:choose>
          <xsl:when test="$stretch='stretch'">
          <a:stretch>
            <a:fillRect />
          </a:stretch>
          </xsl:when>
          <xsl:when test="@style:repeat='stretch'">
          <a:stretch>
            <a:fillRect />
          </a:stretch>
          </xsl:when>
        </xsl:choose>
       
        <xsl:if test="@draw:fill-image-ref-point-x or @draw:fill-image-ref-point-y">
          <a:tile tx="0" ty="0" flip="none">
            <xsl:if test="@draw:fill-image-ref-point-x">
              <xsl:attribute name="sx">
                <xsl:value-of select="number(substring-before(@draw:fill-image-ref-point-x,'%')) * 1000"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@draw:fill-image-ref-point-y">
              <xsl:attribute name="sy">
                <xsl:value-of select="number(substring-before(@draw:fill-image-ref-point-y,'%')) * 1000"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@draw:fill-image-ref-point">
              <xsl:attribute name="algn">
                <xsl:choose>
                  <xsl:when test="@draw:fill-image-ref-point='top-left'">
                    <xsl:value-of select ="'tl'"/>
                  </xsl:when>
                  <xsl:when test="@draw:fill-image-ref-point='top'">
                    <xsl:value-of select ="'t'"/>
                  </xsl:when>
                  <xsl:when test="@draw:fill-image-ref-point='top-right'">
                    <xsl:value-of select ="'tr'"/>
                  </xsl:when>
                  <xsl:when test="@draw:fill-image-ref-point='right'">
                    <xsl:value-of select ="'r'"/>
                  </xsl:when>
                  <xsl:when test="@draw:fill-image-ref-point='bottom-left'">
                    <xsl:value-of select ="'bl'"/>
                  </xsl:when>
                  <xsl:when test="@draw:fill-image-ref-point='bottom-right'">
                    <xsl:value-of select ="'br'"/>
                  </xsl:when>
                  <xsl:when test="@draw:fill-image-ref-point='bottom'">
                    <xsl:value-of select ="'b'"/>
                  </xsl:when>
                  <xsl:when test="@draw:fill-image-ref-point='center'">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </a:tile>
        </xsl:if>
      </a:blipFill>
    </xsl:if>
    <xsl:if test="not(document('styles.xml')/office:document-styles/office:styles/draw:fill-image[@draw:name=$var_imageName])">
      <a:noFill/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template name="tmpshapeTransperancy" >
    <xsl:param name="tranparency"/>
    <xsl:choose>
      <xsl:when test ="$tranparency ='0'">
        <a:alpha val="0"/>
      </xsl:when>
      <xsl:when test="$tranparency !=''">
        <a:alpha>
          <xsl:attribute name="val">
            <xsl:value-of select="format-number($tranparency * 1000,'#')" />
          </xsl:attribute>
        </a:alpha>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpGradientFill">
    <xsl:param name="gradStyleName"/>
    <xsl:param name="opacity"/>
    <a:gradFill flip="none" rotWithShape="0">

                <xsl:for-each select="document('styles.xml')//draw:gradient[@draw:name= $gradStyleName]">
        <a:gsLst>
          <xsl:choose>
            <xsl:when test="@draw:style='linear'">
          <a:gs pos="0">
								<xsl:variable name="varSrgbVal">
                    <xsl:if test="@draw:start-color">
                      <xsl:value-of select="substring-after(@draw:start-color,'#')" />
                    </xsl:if>
                    <xsl:if test="not(@draw:start-color)">
                      <xsl:value-of select="'ffffff'" />
                    </xsl:if>
								</xsl:variable>
								<xsl:if test="$varSrgbVal != ''">
									<a:srgbClr>
										<xsl:attribute name ="val">
											<xsl:value-of select ="$varSrgbVal"/>
                  </xsl:attribute>
               <xsl:call-template name="tmpshapeTransperancy">
                <xsl:with-param name="tranparency" select="$opacity"/>
              </xsl:call-template>
            </a:srgbClr >
								</xsl:if>
          </a:gs>
          <a:gs pos="100000">
								<xsl:variable name="varSrgbVal">
                <xsl:if test="@draw:end-color">
                  <xsl:value-of select="substring-after(@draw:end-color,'#')" />
                </xsl:if>
                <xsl:if test="not(@draw:end-color)">
                  <xsl:value-of select="'ffffff'" />
                </xsl:if>
								</xsl:variable>
								<xsl:if test="$varSrgbVal != ''">
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select ="$varSrgbVal"/>
              </xsl:attribute>
              <xsl:call-template name="tmpshapeTransperancy">
                <xsl:with-param name="tranparency" select="$opacity"/>
              </xsl:call-template>
            </a:srgbClr>
								</xsl:if>
          </a:gs>
           
            </xsl:when>
            <xsl:otherwise>
              <a:gs pos="0">
								<xsl:variable name="varSrgbVal">
                    <xsl:if test="@draw:end-color">
                      <xsl:value-of select="substring-after(@draw:end-color,'#')" />
                    </xsl:if>
                    <xsl:if test="not(@draw:end-color)">
                      <xsl:value-of select="'ffffff'" />
                    </xsl:if>
								</xsl:variable>
								<xsl:if test="$varSrgbVal != ''">
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select ="$varSrgbVal"/>
                  </xsl:attribute>
                  <xsl:call-template name="tmpshapeTransperancy">
                    <xsl:with-param name="tranparency" select="$opacity"/>
                  </xsl:call-template>
                </a:srgbClr>
								</xsl:if>
              </a:gs>
              <a:gs pos="100000">
								<xsl:variable name="varSrgbVal">
                    <xsl:if test="@draw:start-color">
                      <xsl:value-of select="substring-after(@draw:start-color,'#')" />
                    </xsl:if>
                    <xsl:if test="not(@draw:start-color)">
                      <xsl:value-of select="'ffffff'" />
                    </xsl:if>
								</xsl:variable>
								<xsl:if test="$varSrgbVal != ''">
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select ="$varSrgbVal"/>
                  </xsl:attribute>
                  <xsl:call-template name="tmpshapeTransperancy">
                    <xsl:with-param name="tranparency" select="$opacity"/>
                  </xsl:call-template>
                </a:srgbClr >
								</xsl:if>
              </a:gs>
            </xsl:otherwise>
          </xsl:choose>

        </a:gsLst>
           <xsl:choose>
        <xsl:when test="@draw:style='radial' or @draw:style='ellipsoid'">
          <a:path path="circle">
              <xsl:call-template name="tmpFillToRect"/>
            </a:path>
            <!--<xsl:call-template name="tmpTileToRect"/>-->
          </xsl:when>
          <!--<xsl:when test="@draw:style='ellipsoid'">
            <a:path path="shape">
               <xsl:call-template name="tmpFillToRect"/>
          </a:path>
        </xsl:when>-->
        <xsl:when test="@draw:style='linear'">
            <a:lin  scaled="1">
              <xsl:if test="@draw:angle!=''">
                <xsl:attribute name="ang">
                  <xsl:choose>
                    <xsl:when test="@draw:angle">
                      <xsl:variable name="angleValue">
                      <xsl:value-of select="round(((( ( -1 * @draw:angle) + 900 ) mod 3600)  div 10) * 60000)"/>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$angleValue &lt; 0">
                          <xsl:value-of select="-1 * $angleValue "/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$angleValue "/>
                      </xsl:otherwise>
                      </xsl:choose>
                     
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </xsl:if>
            </a:lin>
          <a:tileRect/>
        </xsl:when>
        <xsl:when test="@draw:style='rectangular' or @draw:style='square'">
          <a:path path="rect">
              <xsl:call-template name="tmpFillToRect"/>
          </a:path>
            <!--<xsl:call-template name="tmpTileToRect"/>-->
        </xsl:when>
        <xsl:otherwise>
          <a:lin ang="0" scaled="1"/>
          <a:tileRect/>
        </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </a:gradFill>
  </xsl:template>
  <xsl:template name="tmpFillToRect">
    <a:fillToRect>
      <xsl:if test="@draw:cx">
        <xsl:attribute name="l">
          <xsl:value-of select="substring-before(@draw:cx,'%') * 1000"/>
        </xsl:attribute>
        <xsl:attribute name="r">
          <xsl:value-of select="substring-before(@draw:cx,'%') * 1000"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@draw:cy">
        <xsl:attribute name="t">
          <xsl:value-of select="substring-before(@draw:cy,'%') * 1000"/>
        </xsl:attribute>
        <xsl:attribute name="b">
          <xsl:value-of select="substring-before(@draw:cy,'%') * 1000"/>
        </xsl:attribute>
      </xsl:if>
      </a:fillToRect>
      </xsl:template>
  <xsl:template name="tmpTileToRect">
    <xsl:choose>
      <xsl:when test="@draw:cx and @draw:cy">
        <xsl:choose>
          <xsl:when test="substring-before(@draw:cx,'%') =100 and substring-before(@draw:cy,'%') = 100">
            <a:tileRect/>
          </xsl:when>
          <xsl:when test="substring-before(@draw:cx,'%') =0 and substring-before(@draw:cy,'%') = 0">
            <a:tileRect l="-100000" t="-100000"/>
          </xsl:when>
          <xsl:when test="substring-before(@draw:cx,'%') =100 and substring-before(@draw:cy,'%') = 0">
            <a:tileRect r="-100000" t="-100000"/>
          </xsl:when>
          <xsl:when test="substring-before(@draw:cx,'%') =0 and substring-before(@draw:cy,'%') = 100">
            <a:tileRect l="-100000" b="-100000"/>
          </xsl:when>
          <xsl:when test="substring-before(@draw:cx,'%') > 0 and substring-before(@draw:cy,'%') > 0">
            <a:tileRect/>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <a:tileRect/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="getClassName">
    <xsl:param name ="clsName"/>
    <!-- Node added by vijayeta,to insert font sizes to inner levels-->
    <xsl:param name ="lvl"/>
    <xsl:param name ="masterPageName"/>
    <xsl:choose >
      <xsl:when test ="$clsName='title'">
        <xsl:choose>
          <xsl:when test="$masterPageName='Standard'">
						<xsl:value-of select ="'Standard-title'"/>
					</xsl:when>
					<xsl:when test="$masterPageName='Default'">
						<xsl:value-of select ="'Default-title'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select ="concat($masterPageName,'-title')"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:when test ="$clsName='subtitle'">
				<xsl:choose>
					<xsl:when test="$masterPageName='Standard'">
						<xsl:value-of select ="'Standard-subtitle'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select ="concat($masterPageName,'-subtitle')"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!-- By vijayeta class name s in stylea.xml for differant levels-->
			<xsl:when test ="$clsName='outline'">
				<xsl:choose>
					<xsl:when test="$masterPageName='Standard'">
						<xsl:value-of select ="concat('Standard-outline',$lvl+1)"/>
					</xsl:when>
					<xsl:when test="$masterPageName='Default'">
						<xsl:value-of select ="concat('Default-outline',$lvl+11)"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select ="concat($masterPageName,'-outline',$lvl+1)"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when >
			<xsl:when test="$clsName='standard'">
				<xsl:value-of select ="'standard'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select ="$clsName"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="getDefaultFonaName">
		<xsl:param name ="className"/>
		<xsl:param name ="lvl"/>
		<xsl:param name ="masterPageName" />
		<xsl:variable name ="defaultClsName">
			<xsl:call-template name ="getClassName">
				<xsl:with-param name ="clsName" select="$className"/>
				<!-- Node added by vijayeta,to insert font sizes to inner levels-->
				<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
				<xsl:with-param name ="lvl" select ="$lvl"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:choose >
			<xsl:when test ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:paragraph-properties/@fo:font-family">
				<xsl:variable name ="FontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:paragraph-properties/@fo:font-family"/>
				</xsl:variable>
				<xsl:value-of select ="translate($FontName, &quot;'&quot;,'')" />
			</xsl:when>
			<xsl:when test ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family">
				<xsl:variable name ="FontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family"/>
				</xsl:variable >
				<xsl:value-of select ="translate($FontName, &quot;'&quot;,'')" />
			</xsl:when>
			<!-- Added by lohith - to access default Font family-->
			<xsl:when test ="$defaultClsName='standard'">
				<xsl:variable name ="shapeFontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family"/>
				</xsl:variable>
				<xsl:value-of select ="translate($shapeFontName, &quot;'&quot;,'')" />
			</xsl:when>
			<xsl:when test ="not(document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family)">
				<xsl:variable name ="parentFontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/@style:parent-style-name"/>
				</xsl:variable>
				<xsl:variable name ="shapeFontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $parentFontName]/style:text-properties/@fo:font-family"/>
				</xsl:variable>
				<xsl:if test ="$shapeFontName !=''">
					<xsl:value-of select ="translate($shapeFontName, &quot;'&quot;,'')" />
				</xsl:if>
				<xsl:if test ="$shapeFontName =''">
					<xsl:value-of select ="'Arial'"/>
				</xsl:if>
			</xsl:when>
			<xsl:otherwise >
				<xsl:value-of select ="'Arial'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
   <xsl:template name="tmpgetDefualtTextProp">
    <xsl:param name="parentStyleName"/>
    <xsl:param name="attrName"/>
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:for-each select="document('styles.xml')//style:style[@style:name = $parentStyleName]/style:text-properties">
      <xsl:variable name="prStyleName" select="./parent::node()/@style:parent-style-name"/>
      <xsl:choose>
        <xsl:when test="$attrName='Fontcolor'">
          <xsl:choose>
            <xsl:when test="@fo:color">
              <a:solidFill>
                <xsl:variable name="varSrgbVal">
                  <xsl:value-of select ="translate(substring-after(@fo:color,'#'),$lcletters,$ucletters)"/>
                </xsl:variable>
                <xsl:if test="$varSrgbVal != ''">
                <a:srgbClr  >
                  <xsl:attribute name ="val">
                      <xsl:value-of select ="$varSrgbVal"/>
                  </xsl:attribute>
                </a:srgbClr >
                </xsl:if>
              </a:solidFill>
            </xsl:when>
            <xsl:when test="@style:use-window-font-color='true'">
              <a:solidFill>
                <a:sysClr val="windowText"/>
              </a:solidFill>
            </xsl:when>
            <xsl:when test="not(@fo:color) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Fontcolor'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
          <xsl:when test="$attrName='Fontsize'">
<!--Office 2007 Sp2-->
            <xsl:variable name="fontSize">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length" select="@fo:font-size"/>
              </xsl:call-template>
            </xsl:variable>

          <xsl:choose>
            <xsl:when test="$fontSize > 0">
              <xsl:attribute name ="sz">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'pt'"/>
                  <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="substring-before(@style:font-size-asian,'pt') > 0">
              <xsl:attribute name ="sz">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'pt'"/>
                  <xsl:with-param name ="length" select ="@style:font-size-asian"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="substring-before(@style:font-size-complex,'pt') > 0">
              <xsl:attribute name ="sz">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'pt'"/>
                  <xsl:with-param name ="length" select ="@style:font-size-complex"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="not(@fo:font-size) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Fontsize'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='Fontname'">
          <xsl:choose>
            <xsl:when test="@fo:font-family">
              <a:latin charset="0" >
                <xsl:attribute name ="typeface" >
                  <xsl:value-of select ="translate(@fo:font-family, &quot;'&quot;,'')" />
                </xsl:attribute>
              </a:latin>
            </xsl:when>
            <xsl:when test="not(@fo:font-family) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Fontname'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='Bold'">
          <xsl:choose>
            <xsl:when test="@fo:font-weight[contains(.,'bold')]">
              <xsl:attribute name ="b">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute >
            </xsl:when>
            <xsl:when test="@fo:font-weight[contains(.,'normal')]">
              <xsl:attribute name ="b">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute >
            </xsl:when>
            <xsl:when test="not(@fo:font-weight[contains(.,'bold')]) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Bold'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='italic'">
          <xsl:choose>
            <xsl:when test="@fo:font-style='italic'">
              <xsl:attribute name ="i">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute >
            </xsl:when>
            <xsl:when test="@fo:font-style='normal'">
              <xsl:attribute name ="i">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute >
            </xsl:when >
            <xsl:when test="not(@fo:font-style) and @fo:font-style!='italic' and @fo:font-style!='normal' and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'italic'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='kerning'">
          <xsl:choose>
            <xsl:when test ="@style:letter-kerning = 'true'">
              <xsl:attribute name ="kern">
                <xsl:value-of select="1200"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="@style:letter-kerning = 'false'">
              <xsl:attribute name ="kern">
                <xsl:value-of select="0"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="not(@style:letter-kerning) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'kerning'"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="not(@style:letter-kerning)">
              <xsl:attribute name ="kern">
                <xsl:value-of select="0"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        
        </xsl:when>
        <xsl:when test="$attrName='Textshadow'">
          <xsl:choose>
            <xsl:when test ="@fo:text-shadow != 'none'">
              <a:effectLst>
                <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                  <a:srgbClr val="000000">
                    <a:alpha val="43137" />
                  </a:srgbClr>
                </a:outerShdw>
              </a:effectLst>
            </xsl:when>
            <xsl:when test="not(@fo:text-shadow) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Textshadow'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='strike'">
          <xsl:choose >
            <xsl:when  test="@style:text-line-through-type = 'solid'">
              <xsl:attribute name ="strike">
                <xsl:value-of select ="'sngStrike'"/>
              </xsl:attribute >
            </xsl:when >
            <xsl:when test="@style:text-line-through-type[contains(.,'double')]">
              <xsl:attribute name ="strike">
                <xsl:value-of select ="'dblStrike'"/>
              </xsl:attribute >
            </xsl:when >
            <!-- style:text-line-through-style-->
            <xsl:when test="@style:text-line-through-style = 'solid'">
              <xsl:attribute name ="strike">
                <xsl:value-of select ="'sngStrike'"/>
              </xsl:attribute >
            </xsl:when>
            <xsl:when test="not(@style:text-line-through-type or @style:text-line-through-style) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'strike'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>

        </xsl:when>
        <xsl:when test="$attrName='Underlinecolor'">
          <xsl:choose>
              <xsl:when test="@style:text-underline-color='font-color'">
                <a:uFill>
                  <xsl:call-template name="tmpgetDefualtTextProp">
                    <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
                    <xsl:with-param name="attrName" select="'Fontcolor'"/>
                  </xsl:call-template>
                </a:uFill>
              </xsl:when>
            <xsl:when test ="@style:text-underline-color">
              <a:uFill>
                <a:solidFill>
                  <xsl:variable name="varSrgbVal">
                    <xsl:value-of select ="substring-after(@style:text-underline-color,'#')"/>
                  </xsl:variable>
                  <xsl:if test="$varSrgbVal != ''">
                  <a:srgbClr>
                    <xsl:attribute name ="val">
                        <xsl:value-of select ="$varSrgbVal"/>
                    </xsl:attribute>
                  </a:srgbClr>
                  </xsl:if>                  
                </a:solidFill>
              </a:uFill>
            </xsl:when>
            <xsl:when test="not(@style:text-underline-color) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Underlinecolor'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
   <xsl:template name="tmpgetDefualtParagraphProp">
    <xsl:param name="parentStyleName"/>
    <xsl:param name="attrName"/>
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:for-each select="document('styles.xml')//style:style[@style:name = $parentStyleName]/style:paragraph-properties">
      <xsl:variable name="prStyleName" select="./parent::node()/@style:parent-style-name"/>
      <xsl:choose>
        <xsl:when test="$attrName='Textindent'">
          <xsl:choose>
            <xsl:when test="@fo:text-indent">
              <xsl:variable name ="varIndent">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length">
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="@fo:text-indent"/>
                    </xsl:call-template>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test ="$varIndent!=''">
                <xsl:attribute name ="indent">
                  <xsl:value-of select ="$varIndent"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:when>
            <xsl:when test="not(@fo:text-indent) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtParagraphProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Textindent'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='Textalign'">
          <xsl:choose>
            <xsl:when test="@fo:text-align">
              <xsl:attribute name ="algn">
                <xsl:choose >
                  <xsl:when test ="@fo:text-align='center'">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:when>
                  <xsl:when test ="@fo:text-align='end' or @fo:text-align='right'">
                    <xsl:value-of select ="'r'"/>
                  </xsl:when>
                  <xsl:when test ="@fo:text-align='justify'">
                    <xsl:value-of select ="'just'"/>
                  </xsl:when>
                  <xsl:when test ="@fo:text-align='start'">
                    <xsl:value-of select ="'l'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="not(@fo:text-align) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtParagraphProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Textalign'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='MarginLeft'">
          <xsl:choose>
            <xsl:when test="@fo:margin-left">
              <xsl:variable name ="varMarginLeft">
                <xsl:call-template name ="convertToPoints">
                   <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length">
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="@fo:margin-left"/>
                    </xsl:call-template>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test ="$varMarginLeft!=''">
                <xsl:attribute name ="marL">
                  <xsl:value-of select ="$varMarginLeft"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:when>
            <xsl:when test="not(@fo:margin-left) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'MarginLeft'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='MarginTop'">
          <xsl:choose>
            <xsl:when test="@fo:margin-top">
              <a:spcBef>
                <a:spcPts>
                  <xsl:attribute name ="val">
                    <xsl:call-template name ="convertToPointsLineSpacing">
                      <xsl:with-param name="length"  select ="@fo:margin-top"/>
                      <xsl:with-param name ="unit" select ="'cm'"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </a:spcPts>
              </a:spcBef >
            </xsl:when>
            <xsl:when test="not(@fo:margin-top) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'MarginTop'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='MarginBottom'">
          <xsl:choose>
            <xsl:when test="@fo:margin-bottom">
              <a:spcAft>
                <a:spcPts>
                  <xsl:attribute name ="val">
                    <xsl:call-template name ="convertToPointsLineSpacing">
                      <xsl:with-param name="length"  select ="@fo:margin-bottom"/>
                      <xsl:with-param name ="unit" select ="'cm'"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </a:spcPts>
              </a:spcAft >
            </xsl:when>
            <xsl:when test="not(@fo:margin-bottom) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'MarginBottom'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='Linespacing'">
          <xsl:variable name="lineSpc">
            <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name="length"  select ="@fo:letter-spacing"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="lineSpcAtleast">
            <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name="length"  select ="@style:line-height-at-least"/>
            </xsl:call-template>
          </xsl:variable>
            <xsl:choose>
              <xsl:when test ="@fo:line-height and 
					substring-before(@fo:line-height,'%') &gt; 0 and 
					not(substring-before(@fo:line-height,'%') = 100)">
                <a:lnSpc>
                  <a:spcPct>
                    <xsl:attribute name ="val">
                      <xsl:value-of select ="format-number(substring-before(@fo:line-height,'%')* 1000,'#.##') "/>
                    </xsl:attribute>
                  </a:spcPct>
                </a:lnSpc>
              </xsl:when>
            <!--If the line spacing is in terms of Points,multiply the value with 2835-->
            <xsl:when test ="$lineSpc > 0">
                <a:lnSpc>
                  <a:spcPts>
                    <xsl:attribute name ="val">
                      <xsl:call-template name ="convertToPointsLineSpacing">
                       <xsl:with-param name ="unit" select ="'cm'"/>
                      <xsl:with-param name ="length">
                        <xsl:call-template name="convertUnitsToCm">
                          <xsl:with-param name="length"  select ="$lineSpc"/>
                        </xsl:call-template>
                      </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </a:spcPts>
                </a:lnSpc>
              </xsl:when>
            <xsl:when test ="$lineSpcAtleast > 0 ">
                <a:lnSpc>
                  <a:spcPts>
                    <xsl:attribute name ="val">
                      <xsl:call-template name ="convertToPointsLineSpacing">
                       <xsl:with-param name ="unit" select ="'cm'"/>
                      <xsl:with-param name ="length">
                        <xsl:call-template name="convertUnitsToCm">
                          <xsl:with-param name="length"  select ="$lineSpcAtleast"/>
                        </xsl:call-template>
                      </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </a:spcPts>
                </a:lnSpc>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test="$prStyleName !=''">
                  <xsl:call-template name="tmpgetDefualtParagraphProp">
                    <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                    <xsl:with-param name="attrName" select="'Linespacing'"/>
                  </xsl:call-template>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
	<xsl:template name ="getDefaultFontSize">
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
  <xsl:template name="tmpUnderLineStyle">
    <xsl:param name="parentStyleName"/>
    <xsl:param name="flagPresentationClass"/>
    <xsl:param name="prClassName"/>

    <xsl:choose >
      <!-- Added by lohith for fix - 1744082 - Start-->
      <xsl:when test="style:text-properties/@style:text-underline-type = 'single'">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'sng'"/>
					</xsl:attribute >
				</xsl:when>
      <!-- Fix - 1744082 - End-->
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
      <xsl:otherwise >
        <xsl:if test="$flagPresentationClass='No' or $prClassName='subtitle'">
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name = $parentStyleName]">
          <xsl:call-template name="tmpUnderLineStyle">
          </xsl:call-template>
        </xsl:for-each>
          </xsl:if>
      </xsl:otherwise>
			</xsl:choose>
		
	</xsl:template>
	<xsl:template name ="insertTab">
		<xsl:for-each select ="node()">
			<xsl:choose >
				<xsl:when test ="name()=''">
					<xsl:value-of select ="."/>
				</xsl:when>
				<xsl:when test ="name()='text:tab'">
					<xsl:value-of select ="'&#09;'"/>
				</xsl:when >
				<xsl:when test ="name()='text:s'">
          <pxs:s xmlns:pxs="urn:cleverage:xmlns:post-processings:extra-spaces">
                  <xsl:if test="@text:c">
                    <xsl:attribute name="pxs:c">
                      <xsl:value-of select="@text:c"/>
                    </xsl:attribute>
                  </xsl:if>
                </pxs:s>
					<!--<xsl:call-template name ="insertSpace" >
						<xsl:with-param name ="spaceVal" select ="@text:c"/>
					</xsl:call-template>-->
          
				</xsl:when >
				<xsl:when test =".='' and child::node()">
					<xsl:value-of select ="' '"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select ="."/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
		<xsl:if test ="not(node())">
			<xsl:value-of select ="."/>
		</xsl:if>
	</xsl:template>
		<xsl:template name ="paragraphTabstops">
		<a:tabLst>
			<xsl:for-each select ="style:paragraph-properties/style:tab-stops/style:tab-stop">
				<a:tab >
          <xsl:choose>
            <xsl:when test="@style:position='cm' or contains(@style:position,'NaN')">
              <xsl:attribute name ="pos">
              <xsl:value-of select="'0'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
					<xsl:attribute name ="pos">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length">
                  <xsl:call-template name="convertUnitsToCm">
                    <xsl:with-param name="length">
                      <xsl:choose >
                        <xsl:when test ="./parent::node()/parent::node()/@fo:margin-left">
                          <xsl:variable name="Tabposition">
                            <xsl:call-template name="convertUnitsToCm">
                    <xsl:with-param name="length"  select ="@style:position"/>
                  </xsl:call-template>
                          </xsl:variable>
                          <xsl:variable name="MarL">
                            <xsl:call-template name="convertUnitsToCm">
                              <xsl:with-param name="length" select="./parent::node()/parent::node()/@fo:margin-left"/>
                            </xsl:call-template>
                          </xsl:variable>
                          <xsl:choose>
                            <xsl:when test="$MarL!=''">
                          <xsl:value-of select=" $Tabposition + $MarL"/>
                        </xsl:when>
                        <xsl:otherwise>
                              <xsl:value-of select=" $Tabposition"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select ="@style:position"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>  
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
					</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
					<xsl:attribute name ="algn">
						<xsl:choose >
							<xsl:when test ="@style:type ='center'">
								<xsl:value-of select ="'ctr'"/>
							</xsl:when>
							<xsl:when test ="@style:type ='left'">
								<xsl:value-of select ="'l'"/>
							</xsl:when>
							<xsl:when test ="@style:type ='right'">
								<xsl:value-of select ="'r'"/>
							</xsl:when>
							<xsl:when test ="@style:type ='char'">
								<xsl:value-of select ="'dec'"/>
							</xsl:when>
							<xsl:otherwise >
								<xsl:value-of select ="'l'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</a:tab >
			</xsl:for-each>
		</a:tabLst >
	</xsl:template>
	<!-- Template added by lohith - to get the page id -->
	<xsl:template name="getThePageId">
		<xsl:param name="PageName"/>
		<xsl:variable name="customPage">
			<xsl:for-each select="document('content.xml')/office:document-content/office:body/office:presentation/draw:page">
				<xsl:if test="@draw:name = $PageName">
					<xsl:value-of select="position()" />
				</xsl:if>
			</xsl:for-each>
		</xsl:variable>
		<xsl:if test="$customPage > 0 ">
			<xsl:value-of select="$customPage" />
		</xsl:if>
		<!-- Added for expectional case where #page is used to link the slides - Eg: Uni_animation.odp -->
		<xsl:if test="not($customPage > 0)">
			<xsl:if test="format-number(substring-after($PageName,'page'),'#') > 0 ">
				<xsl:value-of select="substring-after($PageName,'page')" />
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name ="MarginTemplateForSlide">
		<xsl:param name ="level"/>
		<xsl:param name="listId"/>
		<xsl:param name="fileName"/>
		<xsl:for-each select ="document($fileName)//text:list-style[@style:name=$listId]">
			<xsl:choose >
				<xsl:when test ="./child::node()[$level]/style:list-level-properties/@text:space-before">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
						<xsl:with-param name ="length" select ="./child::node()[$level]/style:list-level-properties/@text:space-before"/>
					</xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
				</xsl:when>
				<xsl:otherwise >
          <!--edited by vipul Margin Left cant be 0cm-->
					<xsl:value-of select ="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each >
	</xsl:template>
	<xsl:template name="IndentTemplateForSlide">
		<xsl:param name="level"/>
		<xsl:param name ="listId" />
		<xsl:param name ="fileName"/>
		<xsl:for-each select ="document($fileName)//text:list-style[@style:name=$listId]">
			<xsl:choose >
				<xsl:when test ="./child::node()[$level]/style:list-level-properties/@text:space-before">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
						<xsl:with-param name ="length" select ="./child::node()[$level]/style:list-level-properties/@text:min-label-width"/>
					</xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
				</xsl:when>
				<xsl:otherwise >
          <!--edited by vipul indent cant be 0cm-->
					<xsl:value-of select ="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each >
	</xsl:template >
  <!-- Font styles for bulleted text, added by vijayeta-->
  <xsl:template name ="getTextNodeForFontStyle">
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl" />
    <xsl:param name ="HyperlinksForBullets" />
    <xsl:param name ="masterPageName"/>
    <xsl:param name ="fileName"/>
    <xsl:param name ="flagPresentationClass"/>
    
	  <xsl:choose>
		  <!-- level 0 or no level in pptx,level 1 in odp-->
		  <xsl:when test ="./text:p">
			  <xsl:for-each select ="./text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 1 in pptx,level 2 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 2 in pptx,level 3 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 3 in pptx,level 4 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 4 in pptx,level 5 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 5 in pptx,level 6 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 6 in pptx,level 7 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 7 in pptx,level 8 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 8 in pptx,level 9 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 9 in pptx,level 10 in odp in case ther's 10th level in input odp, for example, DRM.odp
			Since pptx does not support 10th level, any text that is of 10th level in odp, 
			is set back to 9th level in pptx.
			date:14th Aug, '07-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
	  </xsl:choose>
  </xsl:template>
  <!--End,to get paragraph style name for each of the levels in multilevelled list-->
  <xsl:template name ="addFontStyleTextLevels">
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl"/>
    <xsl:param name ="HyperlinksForBullets" />
    <xsl:param name ="masterPageName" />
    <xsl:param name ="fileName"/>
    <xsl:param name ="flagPresentationClass"/>

    <xsl:for-each select ="child::node()[position()]">
      <xsl:if test ="name()='text:span'">
        <xsl:if test ="not(./text:line-break)">
          <xsl:if test ="child::node()">
            <a:r>
              <a:rPr smtClean="0">
                <!--Font Size -->
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="not($textId ='')">
                  <xsl:call-template name ="fontStyles">
                    <xsl:with-param name ="Tid" select ="$textId" />
                    <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    <xsl:with-param name ="lvl" select ="$lvl"/>
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    <xsl:with-param name ="fileName" select ="$fileName"/>
                    <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
                    <xsl:with-param name ="parentStyleName" select ="$prClassName"/>
                  </xsl:call-template>
                </xsl:if>
              <xsl:if test="name()='text:a' or ./text:a">
                  <xsl:copy-of select="$HyperlinksForBullets"/>
                </xsl:if>
              </a:rPr >
              <a:t>
                <xsl:call-template name ="insertTab" />
              </a:t>
            </a:r>
          </xsl:if>
          <!-- Bug 1744106 fixed by vijayeta, date 16th Aug '07, add a new templaet to set font size and family in endPara-->
          <xsl:if test ="not(child::node())">
            <a:endParaRPr dirty="0" smtClean="0" >
              <xsl:if test ="not(@text:style-name ='')">
                <xsl:call-template name ="getFontSizeFamilyFromContentEndPara">
                  <xsl:with-param name ="Tid" select ="@text:style-name"/>
                </xsl:call-template>
              </xsl:if >
              <xsl:if test ="@text:style-name =''">
                <a:endParaRPr dirty="0" smtClean="0"/>
              </xsl:if>
            </a:endParaRPr>
          </xsl:if>
          <!--End, Bug 1744106 fixed by vijayeta, date 16th Aug '07, add a new template to set font size and family in endPara-->
        </xsl:if>
        <xsl:if test ="./text:line-break">
          <xsl:call-template name ="processBR">
            <xsl:with-param name ="T" select ="@text:style-name" />
            <xsl:with-param name ="prClassName" select ="$prClassName"/>
            <xsl:with-param name ="inSideSpan" select ="'Yes'"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:if>
      <xsl:if test ="name()='text:line-break'">
        <xsl:call-template name ="processBR">
          <xsl:with-param name ="T" select ="@text:style-name" />
          <xsl:with-param name ="prClassName" select ="$prClassName"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test ="name()='text:tab'">
        <a:r>
          <a:rPr smtClean="0">
            <!--Font Size -->
            <xsl:variable name ="textId">
              <xsl:value-of select ="@text:style-name"/>
            </xsl:variable>
            <xsl:if test ="not($textId ='')">
              <xsl:call-template name ="fontStyles">
                <xsl:with-param name ="Tid" select ="$textId" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
                <xsl:with-param name ="lvl" select ="$lvl"/>
                <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                <xsl:with-param name ="fileName" select ="$fileName"/>
                <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
                <xsl:with-param name ="parentStyleName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
         
          </a:rPr>
          <a:t>
            <xsl:value-of select ="'&#09;'"/>
          </a:t>
        </a:r>

      </xsl:if >
      <xsl:if test ="not(name()='text:span' or name()='text:line-break' or name()='text:tab') ">
        <a:r>
          <a:rPr smtClean="0">
            <!--Font Size -->
            <xsl:variable name ="textId">
              <xsl:value-of select ="./parent::node()/@text:style-name"/>
            </xsl:variable>
            <xsl:if test ="not($textId ='')">
              <xsl:call-template name ="fontStyles">
                <xsl:with-param name ="Tid" select ="$textId" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
                <xsl:with-param name ="lvl" select ="$lvl"/>
                <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                <xsl:with-param name ="fileName" select ="$fileName"/>
                <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
                <xsl:with-param name ="parentStyleName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test="name()='text:a' or ./text:a">
              <xsl:copy-of select="$HyperlinksForBullets"/>
            </xsl:if>
          </a:rPr >
          <a:t>
            <xsl:call-template name ="insertTab" />
          </a:t>
        </a:r>
      </xsl:if >
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="processBR">
    <xsl:param name ="T" />
    <xsl:param name ="inSideSpan" />
    <xsl:if test="$inSideSpan='Yes'">
      <xsl:if test="./node()">
    <a:r>
      <a:rPr smtClean="0">
        <!--Font Size -->
            <xsl:if test ="not($T ='')">
          <xsl:call-template name ="fontStyles">
            <xsl:with-param name ="Tid" select ="$T" />
          </xsl:call-template>
        </xsl:if>
      </a:rPr >
      <a:t>
        <xsl:call-template name ="insertTab" />
            <!--<xsl:value-of select="./node()"/>-->
      </a:t>
    </a:r>
      </xsl:if>
    </xsl:if>
    <a:br>
      <a:rPr smtClean="0">
        <!--Font Size -->
        <xsl:if test ="not($T ='')">
          <xsl:call-template name ="fontStyles">
            <xsl:with-param name ="Tid" select ="$T" />
          </xsl:call-template>
        </xsl:if>
        <xsl:if test ="$T =''">
          <a:latin charset="0" typeface="Arial" />
        </xsl:if >
      </a:rPr >

    </a:br>
  </xsl:template>
  <!-- Bug 1744106 fixed by vijayeta, date 16th Aug '07, add a new templaet to set font size and family in endPara-->
  <xsl:template name="getFontSizeFamilyFromContentEndPara">
    <xsl:param name="Tid" />
    <xsl:for-each select="document('content.xml')//office:automatic-styles/style:style[@style:name =$Tid ]">
<!--Office 2007 Sp2-->

      <xsl:variable name="fontSize">
        <xsl:call-template name="point-measure">
          <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$fontSize > 0">
        <xsl:attribute name="sz">
          <xsl:call-template name="STTextFontSizeInPoints">
            <xsl:with-param name="unit" select="'pt'" />
            <xsl:with-param name="length" select="concat($fontSize,'pt')" />
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="style:text-properties/@fo:font-family">
        <a:latin charset="0">
          <xsl:attribute name="typeface">
            <!--  fo:font-family-->
            <xsl:value-of select="translate(style:text-properties/@fo:font-family,&quot;'&quot;,'')" />
          </xsl:attribute>
        </a:latin>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="tmpSMfontStyles">
    <xsl:param name ="TextStyleID"/>
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl"/>
    <xsl:param name ="masterPageName"/>
    <xsl:message terminate="no">progress:text:p</xsl:message>

    <xsl:for-each  select ="document('styles.xml')//style:style[@style:name =$TextStyleID ]">
      <xsl:if test="position()=1">
        <xsl:choose>
          <xsl:when test="style:text-properties/@style:language-asian and style:text-properties/@style:country-asian">
            <xsl:attribute name ="lang">
              <xsl:value-of select="concat(style:text-properties/@style:language-asian,'-',style:text-properties/@style:country-asian)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="lang">
              <xsl:value-of select="'en-US'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      <xsl:message terminate="no">progress:text:p</xsl:message>
      <!-- Added by lohith :substring-before(style:text-properties/@fo:font-size,'pt')&gt; 0  because sz(font size) shouldnt be zero - 16filesbug-->
<!--Office 2007 Sp2-->

        <xsl:variable name="fontSize">
          <xsl:call-template name="point-measure">
            <xsl:with-param name ="length" select ="style:text-properties/@fo:font-size"/>
          </xsl:call-template>
        </xsl:variable>
      <xsl:if test="$fontSize &gt; 0 ">
        <xsl:attribute name ="sz">
          <xsl:call-template name ="STTextFontSizeInPoints">
            <xsl:with-param name ="unit" select ="'pt'"/>
            <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <!--Superscript and SubScript for Text added by Mathi on 31st Jul 2007-->
      <xsl:if test="style:text-properties/@style:text-position">
        <xsl:call-template name="tmpSuperSubScriptForward"/>
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
      <xsl:choose>
        <xsl:when test="style:text-properties/@fo:font-style='italic'">
        <xsl:attribute name ="i">
          <xsl:value-of select ="'1'"/>
        </xsl:attribute >
        </xsl:when >
        <xsl:when test="style:text-properties/@fo:font-style='normal'">
          <xsl:attribute name ="i">
            <xsl:value-of select ="'0'"/>
          </xsl:attribute >
        </xsl:when >
      </xsl:choose>
      <!-- Font underline-->
      <xsl:call-template name="tmpUnderLineStyle"/>

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
      <xsl:call-template name="tmpCharacterSpacing"/>
      <!--Color Node set as standard colors -->
      <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
      <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
      <xsl:if test ="style:text-properties/@fo:color">
        <a:solidFill>
            <xsl:variable name="varSrgbVal">
              <xsl:value-of select ="translate(substring-after(style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
            </xsl:variable>
            <xsl:if test="$varSrgbVal != ''">
          <a:srgbClr  >
            <xsl:attribute name ="val">
                  <xsl:value-of select ="$varSrgbVal"/>
            </xsl:attribute>
          </a:srgbClr >
            </xsl:if>
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
              <xsl:variable name="varSrgbVal">
                <xsl:value-of select ="substring-after(style:text-properties/style:text-underline-color,'#')"/>
              </xsl:variable>
              <xsl:if test="$varSrgbVal != ''">
            <a:srgbClr>
              <xsl:attribute name ="val">
                    <xsl:value-of select ="$varSrgbVal"/>
              </xsl:attribute>
            </a:srgbClr>
              </xsl:if>
          </a:solidFill>
        </a:uFill>
      </xsl:if>
      </xsl:if>
    </xsl:for-each >
  </xsl:template>
  <xsl:template name ="tmpSMDefaultfontStyles">
    <xsl:param name ="TextStyleID"/>

    <xsl:for-each  select ="document('styles.xml')//style:style[@style:name =$TextStyleID ]">
      <xsl:if test="position()=1">
        <xsl:choose>
          <xsl:when test="style:text-properties/@style:language-asian and style:text-properties/@style:country-asian">
            <xsl:attribute name ="lang">
              <xsl:value-of select="concat(style:text-properties/@style:language-asian,'-',style:text-properties/@style:country-asian)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="lang">
              <xsl:value-of select="'en-US'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
<!--Office 2007 Sp2-->

        <xsl:variable name="fontSize">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
          </xsl:call-template>
        </xsl:variable>
      <xsl:if test="$fontSize &gt; 0 ">
        <xsl:attribute name ="sz">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'pt'"/>
            <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <!--Color Node set as standard colors -->
      <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
      <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
      <xsl:if test ="style:text-properties/@fo:color">
        <a:solidFill>
            <xsl:variable name="varSrgbVal">
              <xsl:value-of select ="translate(substring-after(style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
            </xsl:variable>
            <xsl:if test="$varSrgbVal != ''">
          <a:srgbClr  >
            <xsl:attribute name ="val">
                  <xsl:value-of select ="$varSrgbVal"/>
            </xsl:attribute>
          </a:srgbClr >
            </xsl:if>          
        </a:solidFill>
      </xsl:if>
      <!-- Text Shadow fix -->
      <xsl:if test ="style:text-properties/@fo:font-family">
        <a:latin charset="0" >
          <xsl:attribute name ="typeface" >
            <!-- fo:font-family-->
            <xsl:value-of select ="translate(style:text-properties/@fo:font-family, &quot;'&quot;,'')" />
          </xsl:attribute>
        </a:latin >
      </xsl:if>
      </xsl:if>
    </xsl:for-each >
  </xsl:template>
  <xsl:template name ="SMParagraphStyles" >
    <!--- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
    <xsl:param name ="paraId" />
    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$paraId]">
      <xsl:if test="position()=1">
        <xsl:if test ="style:paragraph-properties/@fo:text-indent">
          <!--fo:text-indent-->
          <xsl:variable name ="varIndent">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:text-indent"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:if test ="$varIndent!=''">
            <xsl:attribute name ="indent">
              <xsl:value-of select ="$varIndent"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if >
        <xsl:if test ="style:paragraph-properties/@fo:text-align">
          <xsl:attribute name ="algn">
            <!--fo:text-align-->
            <xsl:choose >
              <xsl:when test ="style:paragraph-properties/@fo:text-align='center'">
                <xsl:value-of select ="'ctr'"/>
              </xsl:when>
              <xsl:when test ="style:paragraph-properties/@fo:text-align='end' or style:paragraph-properties/@fo:text-align='right'">
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
        <xsl:if test ="style:paragraph-properties/@fo:margin-left">
          <!--fo:margin-left-->
          <xsl:variable name ="varMarginLeft">
            <xsl:call-template name ="convertToPoints">
               <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-left"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:if test ="$varMarginLeft!=''">
            <xsl:attribute name ="marL">
              <xsl:value-of select ="$varMarginLeft"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if >
        <xsl:if test ="style:paragraph-properties/@fo:margin-right">
          <!-- warn if indent after text-->
          <xsl:message terminate="no">translation.odf2oox.paragraphIndentTypeAfterText</xsl:message>
        </xsl:if>
      <xsl:variable name="lineSpc">
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:letter-spacing"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="lineSpcAtleast">
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name="length"  select ="style:paragraph-properties/@style:line-height-at-least"/>
        </xsl:call-template>
      </xsl:variable>
        <xsl:choose>
          <xsl:when test ="style:paragraph-properties/@fo:line-height and 
					substring-before(style:paragraph-properties/@fo:line-height,'%') &gt; 0 and 
					not(substring-before(style:paragraph-properties/@fo:line-height,'%') = 100)">
            <a:lnSpc>
              <a:spcPct>
                <xsl:attribute name ="val">
                  <xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:line-height,'%')* 1000,'#.##') "/>
                </xsl:attribute>
              </a:spcPct>
            </a:lnSpc>
          </xsl:when>
        <xsl:when test ="$lineSpc > 0">
            <a:lnSpc>
              <a:spcPts>
                <xsl:attribute name ="val">
                  <xsl:call-template name ="convertToPointsLineSpacing">
                   <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length">
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="$lineSpc"/>
                    </xsl:call-template>
                  </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </a:spcPts>
            </a:lnSpc>
          </xsl:when>
        <xsl:when test ="$lineSpcAtleast > 0 ">
            <a:lnSpc>
              <a:spcPts>
                <xsl:attribute name ="val">
                  <xsl:call-template name ="convertToPointsLineSpacing">
                   <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length">
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="$lineSpcAtleast"/>
                    </xsl:call-template>
                  </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </a:spcPts>
            </a:lnSpc>
          </xsl:when>
        </xsl:choose>
        <xsl:if test ="style:paragraph-properties/@fo:margin-top">
          <a:spcBef>
            <a:spcPts>
              <xsl:attribute name ="val">
                <!--fo:margin-top-->
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-top"/>
                  <xsl:with-param name ="unit" select ="'cm'"/>
                </xsl:call-template>
                <!--<xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-top,'cm')* 2835) "/>-->
              </xsl:attribute>
            </a:spcPts>
          </a:spcBef >
        </xsl:if>
        <xsl:if test ="style:paragraph-properties/@fo:margin-bottom">
          <a:spcAft>
            <a:spcPts>
              <xsl:attribute name ="val">
                <!--fo:margin-bottom-->
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-bottom"/>
                  <xsl:with-param name ="unit" select ="'cm'"/>
                </xsl:call-template>
                <!--<xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-bottom,'cm')* 2835) "/>-->
              </xsl:attribute>
            </a:spcPts>
          </a:spcAft>
        </xsl:if >
        <xsl:call-template name ="paragraphTabstops"/>
      </xsl:if>
    </xsl:for-each >
  </xsl:template>
  <xsl:template name="Underline">
    <!-- Font underline-->
    <xsl:param name="uStyle"/>
    <xsl:param name="uWidth"/>
    <xsl:param name="uType"/>

    <xsl:choose >
      <xsl:when test="$uStyle='solid' and
								$uType='double'">
        <xsl:value-of select ="'dbl'"/>
      </xsl:when>
      <xsl:when test="$uStyle='solid' and	$uWidth='bold'">
        <xsl:value-of select ="'heavy'"/>
      </xsl:when>
      <xsl:when test="$uStyle='solid' and $uWidth='auto'">
        <xsl:value-of select ="'sng'"/>
      </xsl:when>
      <!-- Dotted lean and dotted bold under line -->
      <xsl:when test="$uStyle='dotted' and	$uWidth='auto'">
        <xsl:value-of select ="'dotted'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dotted' and	$uWidth='bold'">
        <xsl:value-of select ="'dottedHeavy'"/>
      </xsl:when>
      <!-- Dash lean and dash bold underline -->
      <xsl:when test="$uStyle='dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dash'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dashHeavy'"/>
      </xsl:when>
      <!-- Dash long and dash long bold -->
      <xsl:when test="$uStyle='long-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dashLong'"/>
      </xsl:when>
      <xsl:when test="$uStyle='long-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dashLongHeavy'"/>
      </xsl:when>
      <!-- dot Dash and dot dash bold -->
      <xsl:when test="$uStyle='dot-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dotDashLong'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dot-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dotDashHeavy'"/>
      </xsl:when>
      <!-- dot-dot-dash-->
      <xsl:when test="$uStyle='dot-dot-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dotDotDash'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dot-dot-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dotDotDashHeavy'"/>
      </xsl:when>
      <!-- double Wavy -->
      <xsl:when test="$uStyle='wave' and
								$uType='double'">
        <xsl:value-of select ="'wavyDbl'"/>
      </xsl:when>
      <!-- Wavy and Wavy Heavy-->
      <xsl:when test="$uStyle='wave' and
								$uWidth='auto'">
        <xsl:value-of select ="'wavy'"/>
      </xsl:when>
      <xsl:when test="$uStyle='wave' and
								$uWidth='bold'">
        <xsl:value-of select ="'wavyHeavy'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="convertToPointsLineSpacing">
    <xsl:param name ="unit"/>
    <xsl:param name ="length"/>
    <xsl:message terminate="no">progress:text:p</xsl:message>
    <xsl:variable name="lengthVal">
      <xsl:choose>
        <xsl:when test="contains($length,'cm')">
          <xsl:value-of select="substring-before($length,'cm')"/>
        </xsl:when>
        <xsl:when test="contains($length,'pt')">
          <xsl:value-of select="substring-before($length,'pt')"/>
        </xsl:when>
        <xsl:when test="contains($length,'in')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'in') * 2.54 "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'in')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!--mm to cm -->
        <xsl:when test="contains($length,'mm')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'mm') div 10 "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'mm')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
              <!-- km to cm -->
        <xsl:when test="contains($length,'km')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'km') * 100000  "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'km')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>

        <!-- mi to cm -->
        <xsl:when test="contains($length,'mi')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'mi') * 160934.4"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'mi')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- ft to cm -->
        <xsl:when test="contains($length,'ft')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'ft') * 30.48 "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'ft')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- em to cm -->
        <xsl:when test="contains($length,'em')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round(substring-before($length,'em') * .4233) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'em')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- px to cm -->
        <xsl:when test="contains($length,'px')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round(substring-before($length,'px') div 35.43307) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'px')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- pc to cm -->
        <xsl:when test="contains($length,'pc')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round(substring-before($length,'pc') div 2.362) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'pc')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- ex to cm 1 ex to 6 px-->
        <xsl:when test="contains($length,'ex')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="round((substring-before($length,'ex') div 35.43307)* 6) "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'ex')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <!-- m to cm -->
        <xsl:when test="contains($length,'m')">
          <xsl:choose>
            <xsl:when test ="$unit='cm'" >
              <xsl:value-of select="substring-before($length,'m') * 100 "/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="substring-before($length,'m')" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$length"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test ="contains($lengthVal,'NaN')">
        <xsl:value-of select ="0"/>
      </xsl:when>
      <xsl:when test="$lengthVal='0' or $lengthVal='' or ( ($lengthVal &lt; 0) and ($unit != 'cm')) ">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:when test="not(contains($length,'pt'))">
        <xsl:value-of select="concat(format-number($lengthVal * 2835,'#'),'')"/>
      </xsl:when>
      <xsl:when test ="contains($length,'pt')">
        <xsl:value-of select="concat(format-number($lengthVal,'#') * 100 ,'')"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template >
  <!--End of snippet for Bug 1744106, fixed by vijayeta, date 16th Aug '07, add a new template to set font size and family in endPara-->
  <xsl:template name ="getConvertUnit">
    <xsl:param name ="length"/>
    <xsl:choose>
      <xsl:when test="contains($length,'cm')">
        <xsl:value-of select="'cm'"/>
      </xsl:when>
      <xsl:when test="contains($length,'pt')">
        <xsl:value-of select="'pt'"/>
      </xsl:when>
      <xsl:when test="contains($length,'in')">
        <xsl:value-of select="'in'"/>
      </xsl:when>
      <!--mm to cm -->
      <xsl:when test="contains($length,'mm')">
        <xsl:value-of select="'mm'"/>
      </xsl:when>
      <!-- m to cm -->
      <xsl:when test="contains($length,'m')">
        <xsl:value-of select="'m'"/>
      </xsl:when>
      <!-- km to cm -->
      <xsl:when test="contains($length,'km')">
        <xsl:value-of select="'km'"/>
      </xsl:when>
      <!-- mi to cm -->
      <xsl:when test="contains($length,'mi')">
        <xsl:value-of select="'mi'"/>
      </xsl:when>
      <!-- ft to cm -->
      <xsl:when test="contains($length,'ft')">
        <xsl:value-of select="'ft'"/>
      </xsl:when>
      <!-- em to cm -->
      <xsl:when test="contains($length,'em')">
        <xsl:value-of select="'em'"/>
      </xsl:when>
      <!-- px to cm -->
      <xsl:when test="contains($length,'px')">
        <xsl:value-of select="'px'"/>
      </xsl:when>
      <!-- pc to cm -->
      <xsl:when test="contains($length,'pc')">
        <xsl:value-of select="'pc'"/>
      </xsl:when>
      <!-- ex to cm 1 ex to 6 px-->
      <xsl:when test="contains($length,'ex')">
        <xsl:value-of select="'ex'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template >
  <xsl:template name="tmpgroupingCordinates">
    <xsl:param name="InnerGrp"/>

    <p:grpSpPr>
      <a:xfrm>
        <xsl:variable name="Cordinates">
          <xsl:call-template name="tmpGrpCord"/>
       
        </xsl:variable>
      
        <a:off>
          <xsl:attribute name ="x">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyX@',$Cordinates)"/>
          </xsl:attribute>
          <xsl:attribute name ="y">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyY@',$Cordinates)"/>
          </xsl:attribute>
        </a:off>
        <a:ext>
          <xsl:attribute name ="cx">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyCX@',$Cordinates)"/>
          </xsl:attribute>
          <xsl:attribute name ="cy">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyCY@',$Cordinates)"/>
          </xsl:attribute>
        </a:ext>
        <a:chOff>
          <xsl:attribute name ="x">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyChX@',$Cordinates)"/>
          </xsl:attribute>
          <xsl:attribute name ="y">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyChY@',$Cordinates)"/>
          </xsl:attribute>
        </a:chOff>
        <a:chExt>
          <xsl:attribute name ="cx">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyChCX@',$Cordinates)"/>
          </xsl:attribute>
          <xsl:attribute name ="cy">
            <xsl:value-of select="concat($InnerGrp,'group-svgXYWidthHeight:onlyChCY@',$Cordinates)"/>
          </xsl:attribute>
        </a:chExt>
      </a:xfrm>
    </p:grpSpPr>

  </xsl:template>
  <xsl:template name="tmpGrpCord">
    <xsl:param name="Shapetype"/>
          <xsl:for-each select="node()">
                     <xsl:choose>
                   <xsl:when test="name()='draw:frame'">
                <xsl:choose>
                  <xsl:when test="./draw:object or ./draw:object-ole">
                    <xsl:call-template name="tmpgrpValues"/>
                    </xsl:when>
                  <xsl:when test="./draw:image">
                    <xsl:if test ="contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.emf') or contains(./draw:image/@xlink:href,'.wmf')
                    or contains(./draw:image/@xlink:href,'.jfif') or contains(./draw:image/@xlink:href,'.jpe') 
                    or contains(./draw:image/@xlink:href,'.bmp') or contains(./draw:image/@xlink:href,'.dib') 
                    or contains(./draw:image/@xlink:href,'.rle')
                    or contains(./draw:image/@xlink:href,'.bmz') or contains(./draw:image/@xlink:href,'.gfa') 
                    or contains(./draw:image/@xlink:href,'.emz') or contains(./draw:image/@xlink:href,'.wmz')
                    or contains(./draw:image/@xlink:href,'.pcz')
            or contains(./draw:image/@xlink:href,'.tif') or contains(./draw:image/@xlink:href,'.tiff') 
            or contains(./draw:image/@xlink:href,'.cdr') or contains(./draw:image/@xlink:href,'.cgm') or contains(./draw:image/@xlink:href,'.eps') 
            or contains(./draw:image/@xlink:href,'.pct') or contains(./draw:image/@xlink:href,'.pict') or contains(./draw:image/@xlink:href,'.wpg') 
            or contains(./draw:image/@xlink:href,'.jpeg') or contains(./draw:image/@xlink:href,'.gif') or contains(./draw:image/@xlink:href,'.png')
            or contains(./draw:image/@xlink:href,'.jpg')">
                      <xsl:if test="not(./draw:image/@xlink:href[contains(.,'../')])">
                  <xsl:call-template name="tmpgrpValues">
                  </xsl:call-template>
                      </xsl:if>
                    </xsl:if>
                  </xsl:when>
                <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
              <xsl:call-template name="tmpgrpValues"/>
         
                      </xsl:when>
                  </xsl:choose>
                  </xsl:when>
        <xsl:when test="name()='draw:line' or  name()='draw:connector'">
          <xsl:call-template name="tmpgrpValues">
            <xsl:with-param name="Shapetype" select="'Line'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:circle'or name()='draw:custom-shape'">
          <xsl:call-template name="tmpgrpValues"/>
              </xsl:when>
        <xsl:when test="name()='draw:g'">
          <xsl:call-template name="tmpGrpCord">
            <xsl:with-param name="Shapetype" select="$Shapetype"/>
          </xsl:call-template>
                  </xsl:when>
                </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpGroping">
    <xsl:param name="pos"/>
    <xsl:param name="startPos"/>
    <xsl:param name="pageNo"/>
    <xsl:param name="InnerGrp"/>
    <xsl:param name="fileName"/>
    <xsl:param name="master"/>
    <xsl:param name="UniqueId"/>
    <xsl:param name="vmlPageNo"/>
    
    <p:grpSp>
      <p:nvGrpSpPr>
        <p:cNvPr name="Title 1">
          <xsl:attribute name="name">
            <xsl:value-of select="concat('Group ',$pos+1)"/>
          </xsl:attribute>
          <xsl:attribute name="id">
            <xsl:value-of select="$pos+1"/>
          </xsl:attribute>
        </p:cNvPr>
        <p:cNvGrpSpPr>
          <a:grpSpLocks/>
        </p:cNvGrpSpPr>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <xsl:call-template name="tmpgroupingCordinates">
        <xsl:with-param name="InnerGrp" select="$InnerGrp"/>
      </xsl:call-template>
      <xsl:for-each select="node()">
        <xsl:variable name="var_num_1">
          <xsl:value-of select="position()"/>
        </xsl:variable>
        <xsl:variable name="var_num_2">
          <xsl:number level="any"/>
        </xsl:variable>
    
        <xsl:choose>

          <xsl:when test="name()='draw:frame'">
            <xsl:variable name="var_pos" select="position()"/>
            <xsl:variable name="NvPrId">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
              </xsl:call-template>
            </xsl:variable>
      
                <xsl:choose>
                  <xsl:when test="./draw:object or ./draw:object-ole">
                    <xsl:call-template name="tmpOLEObjects">
                      <xsl:with-param name ="pageNo" select ="$pageNo"/>
                      <xsl:with-param name ="shapeCount" select="$NvPrId" />
                      <xsl:with-param name ="grpFlag" select="'true'" />
                    </xsl:call-template>
                  </xsl:when>
              <xsl:when test="./draw:image">
                <xsl:for-each select="./draw:image">
                  <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                    <xsl:if test="not(./@xlink:href[contains(.,'../')])">
                      <xsl:call-template name="InsertPicture">
                        <xsl:with-param name="imageNo">
                          <xsl:choose>
                            <xsl:when test="$vmlPageNo !=''">
                              <xsl:value-of select="concat('slideMaster',$vmlPageNo)"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="$pageNo"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:with-param>
                        <xsl:with-param name="picNo" select ="$NvPrId"/>
                        <xsl:with-param name="fileName">
                          <xsl:choose>
                            <xsl:when test="$master='1'">
                              <xsl:value-of select="'styles.xml'"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'content.xml'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:with-param>
                        <xsl:with-param name ="grpFlag" select="'true'" />
                        <xsl:with-param name ="master" select="$master" />
                        <xsl:with-param name ="grStyle" >
                          <xsl:choose>
                            <xsl:when test="./parent::node()/@draw:style-name">
                              <xsl:value-of select ="./parent::node()/@draw:style-name"/>
                            </xsl:when>
                            <xsl:when test="./parent::node()/@presentation:style-name">
                              <xsl:value-of select ="./parent::node()/@presentation:style-name"/>
                            </xsl:when>
                          </xsl:choose>
                        </xsl:with-param >
                      </xsl:call-template>
                    </xsl:if>
                  </xsl:if>
                </xsl:for-each>
                  </xsl:when>
              <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                <xsl:call-template name ="CreateShape">
                  <xsl:with-param name ="fileName" select ="$fileName"/>
                  <xsl:with-param name ="shapeName" select="'TextBox '" />
                  <xsl:with-param name ="shapeCount" select="$NvPrId" />
                  <xsl:with-param name ="grpFlag" select="'true'" />
                  <xsl:with-param name="UniqueId" select="generate-id()"/>
                </xsl:call-template>
                  </xsl:when>
               </xsl:choose>
              </xsl:when>
          <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:line' or  name()='draw:connector'
          or name()='draw:custom-shape'  or name()='draw:circle'">
            <xsl:variable name="var_pos" select="position()"/>
            <xsl:variable name="NvPrId">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:call-template name ="shapes" >
              <xsl:with-param name ="fileName" select ="$fileName"/>
              <xsl:with-param name ="var_pos" select="$NvPrId" />
              <xsl:with-param name ="grpFlag" select="'true'" />
            </xsl:call-template >
          </xsl:when>
          <xsl:when test="name()='draw:g'">
            <xsl:variable name="var_pos" select="position()"/>
            <xsl:variable name="NvPrId">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select=".">
              <xsl:call-template name="tmpGroping">
                <xsl:with-param name ="pos" select ="$NvPrId"/>
                <xsl:with-param name ="startPos" select ="$startPos+1"/>
                <xsl:with-param name ="pageNo" select ="$pageNo"/>
                <xsl:with-param name ="InnerGrp" select ="'InnerGroup'"/>
                <xsl:with-param name="fileName" select="$fileName"/>
                <xsl:with-param name="master" select="$master"/>
                <xsl:with-param name="vmlPageNo" select="$vmlPageNo"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
            </xsl:choose>
          </xsl:for-each>
    </p:grpSp>
  </xsl:template>
  <xsl:template name="tmpGroupingRelation">
    <xsl:param name="pos"/>
    <xsl:param name="slideNo"/>
    <xsl:param name="FileName"/>
 <xsl:param name="startPos"/>
          <xsl:for-each select="node()">
      <xsl:variable name="var_num_1">
        <xsl:value-of select="position()"/>
      </xsl:variable>
      <xsl:variable name="var_num_2">
        <xsl:number level="any"/>
      </xsl:variable>
            <xsl:choose>
             <xsl:when test="name()='draw:frame'">
          <xsl:variable name="var_pos" select="position()"/>
          <xsl:variable name="Uid" select="generate-id()"/>
              <xsl:variable name="NvPrId">
                <xsl:call-template name="getShapePosTemp">
                  <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
                </xsl:call-template>
              </xsl:variable>
          <xsl:for-each select=".">
                <xsl:choose>
               <xsl:when test="./draw:object or ./draw:object-ole">
                    <xsl:choose>
                      <xsl:when test="document(concat(substring-after(./child::node()[1]/@xlink:href,'./'),'/content.xml'))/child::node() or
                                      document(concat(translate(./child::node()[1]/@xlink:href,'/',''),'/content.xml'))/child::node() ">
                        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                                   xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

                          <xsl:attribute name="Id">
                            <xsl:value-of select="concat('oleObjectImage_',generate-id())"/>
                          </xsl:attribute>
                          <xsl:attribute name="Target">
                            <xsl:value-of select="concat('../media/','oleObjectImage_',generate-id(),'.png')"/>
                          </xsl:attribute>

                        </Relationship>
                      </xsl:when>
                 <xsl:otherwise>
                    <xsl:call-template name="tmpOLEObjectsRel">
                      <xsl:with-param name="slideNo" select="$slideNo"/>
                      <xsl:with-param name ="grpBln" select ="'true'"/>
                    </xsl:call-template>
                  </xsl:otherwise>
                    </xsl:choose>
              </xsl:when>
                <xsl:when test="./draw:image">
                <xsl:for-each select="./draw:image">
                  <xsl:variable name="UniqueID">
                    <xsl:value-of select="generate-id()"/>
                  </xsl:variable>
                  <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                    <xsl:if test="not(./@xlink:href[contains(.,'../')]) and ./@xlink:href != ''">
                      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                        <xsl:attribute name="Id">
                          <xsl:value-of select="concat('slgrpImage',$slideNo,'-',$NvPrId,$UniqueID)" />
                        </xsl:attribute>
                        <xsl:attribute name="Type">
                          <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
                        </xsl:attribute>
                        <xsl:attribute name="Target">
                          <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
                        </xsl:attribute>
                      </Relationship>
                      <!--added by chhavi for picture hyperlink relationship-->
                      <xsl:if test="./following-sibling::node()[name() = 'office:event-listeners']">
                        <xsl:for-each select ="./parent::node()">
                          <xsl:call-template name="tmpOfficeListnerRelationship">
                            <xsl:with-param name="ShapeType" select="'picture'"/>
                            <xsl:with-param name="PostionCount" select="generate-id()"/>
                            <xsl:with-param name="Type" select="'FRAME'"/>
                              </xsl:call-template>
                                    </xsl:for-each>
                      </xsl:if>
                      <!--end here-->
                    </xsl:if>
                  </xsl:if>
                </xsl:for-each>
                          </xsl:when>
            </xsl:choose>
            <xsl:variable name="shapeId">
              <xsl:value-of select="concat('text-box',$NvPrId)"/>
            </xsl:variable>
            <xsl:for-each select="./draw:text-box">
            <xsl:call-template name="tmpHyperLnkBuImgRel">
              <xsl:with-param name ="var_pos" select="$NvPrId" />
              <xsl:with-param name ="shapeId" select="$shapeId" />
              <xsl:with-param name ="grpFlag" select="'true'" />
              <xsl:with-param name ="UniqueId" select="$Uid" />
              
            </xsl:call-template>
            </xsl:for-each>
              <xsl:call-template name="tmpBitmapFillRel">
                <xsl:with-param name ="UniqueId" select="$Uid" />
                <xsl:with-param name ="FileName" select="$FileName" />
                <xsl:with-param name ="prefix" select="'grpbitmap'" />
              </xsl:call-template>
                      </xsl:for-each>
        </xsl:when>
        <xsl:when test="name()='draw:custom-shape'">
          <xsl:variable name="var_pos" select="position()"/>
          <xsl:variable name="NvPrId">
            <xsl:call-template name="getShapePosTemp">
              <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="Uid" select="generate-id()"/>
          <xsl:for-each select=".">
            <xsl:call-template name="tmpBitmapFillRel">
              <xsl:with-param name ="UniqueId" select="$Uid" />
              <xsl:with-param name ="FileName" select="$FileName" />
              <xsl:with-param name ="prefix" select="'grpbitmap'" />
            </xsl:call-template>
            <xsl:variable name="shapeId">
              <xsl:value-of select="concat('custom-shape',$var_pos)"/>
            </xsl:variable>
            <xsl:call-template name="tmpHyperLnkBuImgRel">
              <xsl:with-param name ="var_pos" select="$NvPrId" />
              <xsl:with-param name ="shapeId" select="$shapeId" />
              <xsl:with-param name ="grpFlag" select="'true'" />
              <xsl:with-param name ="UniqueId" select="$Uid" />
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="name()='draw:rect'">
          <xsl:variable name="var_pos" select="position()"/>
          <xsl:variable name="NvPrId">
            <xsl:call-template name="getShapePosTemp">
              <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="Uid" select="generate-id()"/>
          <xsl:for-each select=".">
            <xsl:call-template name="tmpBitmapFillRel">
              <xsl:with-param name ="UniqueId" select="$Uid" />
              <xsl:with-param name ="FileName" select="$FileName" />
              <xsl:with-param name ="prefix" select="'grpbitmap'" />
            </xsl:call-template>
            <xsl:variable name="shapeId">
              <xsl:value-of select="concat('rect',$var_pos)"/>
            </xsl:variable>
            <xsl:call-template name="tmpHyperLnkBuImgRel">
              <xsl:with-param name ="var_pos" select="$NvPrId" />
              <xsl:with-param name ="shapeId" select="$shapeId" />
              <xsl:with-param name ="grpFlag" select="'True'" />
              <xsl:with-param name ="UniqueId" select="$Uid" />
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="name()='draw:ellipse' or name()='draw:circle'">
          <xsl:variable name="var_pos" select="position()"/>
          <xsl:variable name="NvPrId">
            <xsl:call-template name="getShapePosTemp">
              <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="Uid" select="generate-id()"/>
          <xsl:for-each select=".">
            <xsl:call-template name="tmpBitmapFillRel">
              <xsl:with-param name ="UniqueId" select="$Uid" />
              <xsl:with-param name ="FileName" select="$FileName" />
              <xsl:with-param name ="prefix" select="'grpbitmap'" />
            </xsl:call-template>
            <xsl:variable name="shapeId">
              <xsl:value-of select="concat('ellipse',$var_pos)"/>
            </xsl:variable>
            <xsl:call-template name="tmpHyperLnkBuImgRel">
              <xsl:with-param name ="var_pos" select="$NvPrId" />
              <xsl:with-param name ="shapeId" select="$shapeId" />
              <xsl:with-param name ="grpFlag" select="'true'" />
              <xsl:with-param name ="UniqueId" select="$Uid" />
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="name()='draw:g'">
          <xsl:variable name="var_pos" select="position()"/>
          <xsl:variable name="NvPrId">
            <xsl:call-template name="getShapePosTemp">
              <xsl:with-param name="var_pos" select="$pos + $var_pos"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:call-template name="tmpGroupingRelation">
            <xsl:with-param name="slideNo" select="$slideNo"/>
            <xsl:with-param name ="pos" select ="$NvPrId"/>
            <xsl:with-param name="startPos" select="$startPos+1"/>
            <xsl:with-param name ="FileName" select="$FileName" />
          </xsl:call-template>
                          </xsl:when>
                           </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpOLEObjectsRel">
    <xsl:param name ="slideNo"/>
    <xsl:param name ="grpBln"/>
    <xsl:for-each select="./child::node()[1]">
      <xsl:variable name="extension">
        <xsl:variable name="objectName">
          <xsl:choose>
            <xsl:when test="starts-with(@xlink:href,'./')">
              <xsl:value-of select="substring-after(@xlink:href,'./')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@xlink:href"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="substring-after($objectName,'.')!=''">
            <xsl:value-of select="concat('.',substring-after($objectName,'.'))"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="@xlink:href !=''">
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
                                    xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:attribute name="Id">
          <xsl:choose>
            <xsl:when test="$grpBln='true'">
              <xsl:choose>
                <xsl:when test="contains($slideNo,'slideMaster')">
                  <xsl:value-of select="concat('Slidegrp',substring-after($slideNo,'slideMaster'),'_Ole',generate-id())"/>
                </xsl:when>
                <xsl:otherwise>
              <xsl:value-of select="concat('Slidegrp',$slideNo,'_Ole',generate-id())"/>
                </xsl:otherwise>
              </xsl:choose>
              
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat('Slide',$slideNo,'_Ole',generate-id())"/>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:attribute>
        <xsl:choose>
          <xsl:when test="starts-with(@xlink:href,'./')">
            <xsl:variable name="target">
              <xsl:choose>
                <xsl:when test="$extension!=''">
                  <xsl:value-of select="concat('../embeddings/',translate(substring-after(@xlink:href,'./'),' ',''))"/>
                </xsl:when>
                <xsl:when test="$extension=''">
                  <xsl:value-of select="concat('../embeddings/',translate(substring-after(@xlink:href,'./'),' ',''),'.bin')"/>
                </xsl:when>
              </xsl:choose>
            </xsl:variable>
            <xsl:attribute name="Target">
              <xsl:value-of select="$target"/>
              <!--<xsl:value-of select="concat('../embeddings/','oleObject_',generate-id(),$extension)"/-->
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="starts-with(@xlink:href,'//')">
            <xsl:attribute name="Target">
              <xsl:value-of select="concat('file:///\\',translate(substring-after(@xlink:href,'//'),'/','\'))"/>
            </xsl:attribute>
            <xsl:attribute name="TargetMode">
              <xsl:value-of select="'External'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="starts-with(@xlink:href,'/') or starts-with(@xlink:href,'file:///')">
            <xsl:attribute name="Target">
              <xsl:value-of select="concat('file:///',translate(substring-after(@xlink:href,'/'),'/','\'))"/>
            </xsl:attribute>
            <xsl:attribute name="TargetMode">
              <xsl:value-of select="'External'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="starts-with(@xlink:href,'../')">
            <xsl:attribute name="Target">
             <xsl:value-of select ="concat('hyperlink-path:',@xlink:href)"/>
            </xsl:attribute>
            <xsl:attribute name="TargetMode">
              <xsl:value-of select="'External'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="target">
              <xsl:choose>
                <xsl:when test="$extension!=''">
                  <xsl:value-of select="concat('../embeddings/',translate(@xlink:href,' ',''))"/>
                </xsl:when>
                <xsl:when test="$extension=''">
                  <xsl:value-of select="concat('../embeddings/',translate(translate(@xlink:href,' ',''),'/',''),'.bin')"/>
                </xsl:when>
              </xsl:choose>
            </xsl:variable>
            <xsl:attribute name="Target">
              <xsl:value-of select="$target"/>
              <!--<xsl:value-of select="concat('../embeddings/','oleObject_',generate-id(),$extension)"/-->
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </Relationship>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpBulletImageRel">
    <xsl:param name ="var_pos" />
    <xsl:param name ="shapeId" />
    <xsl:param name ="listId" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
    <xsl:param name ="listItemCount" />
    
    <xsl:variable name="forCount" select="position()" />
    <!--to avoid duplicate rIds in .rels file-->
    <xsl:for-each select ="child::node()[position()=1]">
      <xsl:choose >
        <xsl:when test ="name()='text:list-item'">
          <xsl:variable name ="blvl">
            <xsl:call-template name ="getListLevelForTextBox">
              <xsl:with-param name ="levelCount"/>
            </xsl:call-template>
          </xsl:variable >
          <xsl:variable name ="BuImgRel" select ="concat($var_pos,$forCount,$UniqueId)"/>
  
          <xsl:variable name="xhrefValue">
            <xsl:call-template name ="getTextHyperlinksForBulltesForTextBox">
              <xsl:with-param name ="blvl" select="$blvl"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="paragraphId" >
            <xsl:call-template name ="getParaStyleNameForTextBox">
              <xsl:with-param name ="lvl" select ="$blvl"/>
            </xsl:call-template>
          </xsl:variable>
       
          <xsl:variable name ="isNumberingEnabled">
            <xsl:choose >
              <xsl:when test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'true'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:if test="string-length($xhrefValue) > 0">
            <xsl:call-template name="tmpShapeBulletOfficeListnerRel">
              <xsl:with-param name="shapeId" select="$shapeId"/>
              <xsl:with-param name="blvl" select="$blvl"/>
              <xsl:with-param name="xhrefValue" select="$xhrefValue"/>
              <xsl:with-param name="listItemCount" select="$listItemCount"/>
            </xsl:call-template>
          </xsl:if>
          <xsl:for-each select ="document('content.xml')//text:list-style[@style:name=$listId]">
            <xsl:if test ="text:list-level-style-image[@text:level=$blvl+1] and $isNumberingEnabled='true' and text:list-level-style-image[@text:level=$blvl+1]/@xlink:href">
              <!--<xsl:variable name ="rId" select ="concat('buImage',$listId,$blvl+1,$forCount,$PostionCount)"/>-->
              <xsl:variable name ="rId" select ="concat('buImage',$grpFlag,$listId,$BuImgRel,generate-id())"/>
              <xsl:variable name ="imageName">
                <xsl:choose>
                  <xsl:when test="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'Pictures/') != ''">
                    <xsl:value-of select="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'Pictures/')"/>
                  </xsl:when>
                  <xsl:when test="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'media/') != ''">
                    <xsl:value-of select="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'media/')"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:variable>

              <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <xsl:attribute name ="Id">
                  <xsl:value-of  select ="$rId"/>
                </xsl:attribute>
                <xsl:attribute name ="Type" >
                  <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
                </xsl:attribute>
                <xsl:attribute name ="Target">
                  <xsl:value-of select ="concat('../media/',$imageName)"/>
                </xsl:attribute>
              </Relationship >
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpBitmapFillRel">
    <xsl:param name="UniqueId"/>
    <xsl:param name="FileName"/>
    <xsl:param name="prefix"/>
   
    <xsl:variable name ="grStyle" >
      <xsl:choose>
        <xsl:when test="@draw:style-name">
          <xsl:value-of select ="@draw:style-name"/>
        </xsl:when>
        <xsl:when test="@presentation:style-name">
          <xsl:value-of select ="@presentation:style-name"/>
        </xsl:when>
        <xsl:when test="@table:style-name">
          <xsl:value-of select ="@table:style-name"/>
        </xsl:when>
        <xsl:when test="../@table:default-cell-style-name">
          <xsl:value-of select ="../@table:default-cell-style-name"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:for-each select="document($FileName)//style:style[@style:name=$grStyle]/style:graphic-properties">
      <xsl:if test="position()=1">
      <xsl:if test="@draw:fill='bitmap'">
      <xsl:call-template name="tmpBitmapRelationship">
        <xsl:with-param name="UniqueId" select="$UniqueId"/>
        <xsl:with-param name="FileName" select="$FileName"/>
        <xsl:with-param name="prefix" select="$prefix"/>
      </xsl:call-template>
      </xsl:if>
      <xsl:if test="not(@draw:fill) and ./parent::node()/@style:parent-style-name!=''">
        <xsl:variable name="graphicStyle" select="./parent::node()/@style:parent-style-name"/>
        <xsl:for-each select="document('styles.xml')//style:style[@style:name=$graphicStyle]/style:graphic-properties">
            <xsl:if test="position()=1">
          <xsl:if test="@draw:fill='bitmap'">
            <xsl:call-template name="tmpBitmapRelationship">
              <xsl:with-param name="UniqueId" select="$UniqueId"/>
              <xsl:with-param name="FileName" select="$FileName"/>
              <xsl:with-param name="prefix" select="$prefix"/>
            </xsl:call-template>
          </xsl:if>
            </xsl:if>
        </xsl:for-each>
      </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpBitmapRelationship">
    <xsl:param name="UniqueId"/>
    <xsl:param name="FileName"/>
    <xsl:param name="prefix"/>
        <xsl:variable name="var_imageName" select="@draw:fill-image-name"/>
        <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/draw:fill-image[@draw:name=$var_imageName]">
          <xsl:if test="position()=1">
          <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
            <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <xsl:attribute name="Id">
                <xsl:value-of select="concat($prefix,$UniqueId)" />
              </xsl:attribute>
              <xsl:attribute name="Type">
                <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
              </xsl:attribute>
              <xsl:attribute name="Target">
                <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
              </xsl:attribute>
            </Relationship>
          </xsl:if>
          </xsl:if>
        </xsl:for-each>
     </xsl:template>
  <xsl:template name="tmpgrpValues">
    <xsl:param name="shape_count"/>
    <xsl:param name="Shapetype"/>
    <xsl:choose>
      <xsl:when test="$Shapetype='Line'">
        <xsl:variable name="x1">
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length"  select ="@svg:x1"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="x2">
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length"  select ="@svg:x2"/>
          </xsl:call-template>
        </xsl:variable>
         <xsl:variable name="y1">
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length"  select ="@svg:y1"/>
          </xsl:call-template>
        </xsl:variable>
         <xsl:variable name="y2">
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length"  select ="@svg:y2"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="Width">
          <xsl:choose>
            <xsl:when test="number($x1) &lt; number($x2)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="number($x2 - $x1)"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="number($x2) &lt; number($x1)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="number($x1 - $x2)"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$x2 = $x1">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="'0'"/>
              </xsl:call-template>
                  </xsl:when>
               </xsl:choose>
        </xsl:variable>
        <xsl:variable name="Height">
                <xsl:choose>
            <xsl:when test="number($y1) &lt; number($y2)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="number($y2 - $y1)"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="number($y2) &lt; number($y1)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="number($y1 - $y2)"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$y2 = $y1">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="'0'"/>
              </xsl:call-template>
                  </xsl:when>
                 </xsl:choose>
                           </xsl:variable>
        <xsl:variable name="X">
          <xsl:choose>
            <xsl:when test="number($x1) &lt; number($x2)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:x1"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="number($x2) &lt; number($x1)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:x2"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$x2 = $x1">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:x1"/>
              </xsl:call-template>
                  </xsl:when>
                </xsl:choose>

        </xsl:variable>
        <xsl:variable name="Y">
          <xsl:choose>
            <xsl:when test="number($y1) &lt; number($y2)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:y1"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="number($y2) &lt; number($y1)">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:y2"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$y2 = $y1">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:y1"/>
              </xsl:call-template>
              </xsl:when>
            </xsl:choose>


             </xsl:variable>
        <xsl:variable name="Rot">
            <xsl:choose>
              <xsl:when test="@draw:transform">
              <xsl:variable name ="var_rot">
                <xsl:call-template name="tmpGetGroupRotXYVals">
                  <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
                  <xsl:with-param name="XY" select="'ROT'"/>
                </xsl:call-template>
                            </xsl:variable>
              <xsl:value-of select="$var_rot"/>
                          </xsl:when>
                          <xsl:otherwise>
              <xsl:value-of select="'0'"/>
                          </xsl:otherwise>
                        </xsl:choose>
        </xsl:variable>
        <!--<xsl:value-of select="concat('shape',$shape_count,'-',$X,$Y,$Width,$Height,'@')"/>-->
        <xsl:value-of select="concat($X,':',$Y,':',$Width,':',$Height,':','0','@')"/>
                  </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="convertUnit">
          <xsl:call-template name="getConvertUnit">
            <xsl:with-param name="length" select="@draw:transform"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="Width">
                    <xsl:choose>
            <xsl:when test="@svg:width =''">
              <xsl:value-of select="'0'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:width"/>
              </xsl:call-template>
               </xsl:otherwise>
                    </xsl:choose>

        </xsl:variable>
        <xsl:variable name="Height">
                <xsl:choose>
            <xsl:when test="@svg:height =''">
              <xsl:value-of select="'0'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@svg:height"/>
              </xsl:call-template>
                </xsl:otherwise>
                </xsl:choose>

        </xsl:variable>
        <xsl:variable name="X">
                <xsl:choose>
                  <xsl:when test="@svg:x">
                    <xsl:choose>
                      <xsl:when test="@svg:x =''">
                        <xsl:value-of select="'0'"/>
                      </xsl:when>
                      <xsl:otherwise>
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="@svg:x"/>
                    </xsl:call-template>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test="@draw:transform">
                    <xsl:variable name ="x">
                      <xsl:variable name="tmp_X">
                        <xsl:call-template name="tmpGetGroupRotXYVals">
                          <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
                          <xsl:with-param name="unit" select="$convertUnit"/>
                          <xsl:with-param name="XY" select="'X'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:call-template name="convertUnitsToCm">
                        <xsl:with-param name="length" select="concat($tmp_X,$convertUnit)" />
                        </xsl:call-template>
                    </xsl:variable>
                    <xsl:if test="$x=''">
                      <xsl:value-of select="'0'"/>
                    </xsl:if>
                    <xsl:if test="$x!=''">
              <xsl:value-of select="$x"/>
                    </xsl:if>
                
                  </xsl:when>
                  <xsl:otherwise>
              <xsl:value-of select="'0'"/>
                  </xsl:otherwise>
                </xsl:choose>
             </xsl:variable>
        <xsl:variable name="Y">
                        <xsl:choose>
                          <xsl:when test="@svg:y">
                                  <xsl:choose>
                                <xsl:when test="@svg:y =''">
                                  <xsl:value-of select="'0'"/>
                                </xsl:when>
                                <xsl:otherwise>
                            <xsl:call-template name="convertUnitsToCm">
                              <xsl:with-param name="length"  select ="@svg:y"/>
                            </xsl:call-template>
                                </xsl:otherwise>
                              </xsl:choose>

                         
                            
                          </xsl:when>
                          <xsl:when test="@draw:transform">
                            <xsl:variable name ="y">
                              <xsl:variable name="tmp_Y">
                                <xsl:call-template name="tmpGetGroupRotXYVals">
                                  <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
                                  <xsl:with-param name="unit" select="$convertUnit"/>
                                  <xsl:with-param name="XY" select="'Y'"/>
                                </xsl:call-template>
                              </xsl:variable>
                              <xsl:call-template name="convertUnitsToCm">
                                <xsl:with-param name="length" select="concat($tmp_Y,$convertUnit)" />
                              </xsl:call-template>
                            </xsl:variable>
                            <xsl:if test="$y=''">
                              <xsl:value-of select="'0'"/>
                            </xsl:if>
                            <xsl:if test="$y!=''">
              <xsl:value-of select="$y"/>
                            </xsl:if>
                          </xsl:when>
                          <xsl:otherwise>
              <xsl:value-of select="'0'"/>
                          </xsl:otherwise>
                        </xsl:choose>
        </xsl:variable>
        <xsl:variable name="Rot">
                    <xsl:choose>
                      <xsl:when test="@draw:transform">
              <xsl:variable name ="var_rot">
                <xsl:call-template name="tmpGetGroupRotXYVals">
                  <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
                  <xsl:with-param name="XY" select="'ROT'"/>
                </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test="$var_rot=''">
                          <xsl:value-of select="'0'"/>
                        </xsl:if>
                        <xsl:if test="$var_rot!=''">
              <xsl:value-of select="$var_rot"/>
                        </xsl:if>
                      </xsl:when>
                      <xsl:otherwise>
              <xsl:value-of select="'0'"/>
                      </xsl:otherwise>
                    </xsl:choose>
        </xsl:variable>
        <!--<xsl:value-of select="concat('shape',$shape_count,'-',$X,$Y,$Width,$Height,'@')"/>-->
          <xsl:value-of select="concat($X,':',$Y,':',$Width,':',$Height,':',$Rot,'@')"/>
      </xsl:otherwise>
                </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpGetGroupRotXYVals">
    <xsl:param name="drawTranformVal"/>
    <xsl:param name="unit"/>
    <xsl:param name="XY"/>
    
    <xsl:variable name="tmp_XY">
      <xsl:choose>
        <xsl:when test="$drawTranformVal!=''">
      <xsl:value-of select="substring-after(@draw:transform,'translate')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
     
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$XY='X'">
        <xsl:choose>
          <xsl:when test="contains($tmp_XY,'translate')">
            <!--fix for SP2 shape rotation issue-->
            <xsl:variable name="x1">
              <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),$unit)" />
            </xsl:variable>
            <xsl:variable name="x2">
            <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after($tmp_XY,'translate'),'('),')'),' '),$unit)" />
            </xsl:variable>
            <xsl:value-of select="number($x1) + number($x2) "/>
          </xsl:when>
          <xsl:when test="$drawTranformVal!=''">
            <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after($drawTranformVal,'translate'),'('),')'),' '),$unit)" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$XY='Y'">
        <xsl:choose>
          <xsl:when test="contains($tmp_XY,'translate')">
            <!--fix for SP2 shape rotation issue-->
            <xsl:variable name="y1">
             <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),$unit)" />
            </xsl:variable>
            <xsl:variable name="y2">
            <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after($tmp_XY,'translate'),'('),')'),' '),$unit)" />
            </xsl:variable>
            <xsl:value-of select="number($y1) + number($y2) "/>
          </xsl:when>
          <xsl:when test="$drawTranformVal!=''">
            <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after($drawTranformVal,'translate'),'('),')'),' '),$unit)" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$XY='ROT'">
        <xsl:choose>
          <xsl:when test="$drawTranformVal!=''">
        <xsl:value-of select="substring-before(substring-after(substring-after(@draw:transform,'rotate'),'('),')')" />
        <!--<xsl:value-of select="substring-after(substring-before(substring-before($drawTranformVal,'translate'),')'),'(')" />-->
      </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
        
      </xsl:when>
    </xsl:choose>


  </xsl:template>
  <xsl:template name ="tmpdrawCordinates">
    <xsl:param name="OLETAB"/>
    <xsl:param name="grpFlag"/>
    <xsl:choose>
      <xsl:when test="$OLETAB='true'">
        <p:xfrm>
          <xsl:choose>
            <xsl:when test="$grpFlag='true'">
              <xsl:call-template name ="tmpGroupdrawCordinates"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="tmpGetShapeCordinates"/>
              </xsl:otherwise>
          </xsl:choose>
        </p:xfrm>
      </xsl:when>
      <xsl:otherwise>
        <a:xfrm>
          <xsl:call-template name="tmpGetShapeCordinates">
          </xsl:call-template>
        </a:xfrm>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpGetShapeCordinates">
    <xsl:param name="grpFlag"/>
    <xsl:variable name="convertUnit">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="@draw:transform"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="width" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:width"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="height" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:height"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="x" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:x"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="y" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:y"/>
      </xsl:call-template>
    </xsl:variable>
        <xsl:variable name ="angle">
          <!--<xsl:value-of select="substring-after(substring-before(substring-before(@draw:transform,'translate'),')'),'(')" />-->
          <xsl:call-template name="tmpGetGroupRotXYVals">
            <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
            <xsl:with-param name="XY" select="'ROT'"/>
          </xsl:call-template>
      </xsl:variable>
      <xsl:variable name ="x2">
        <xsl:variable name="tmp_X">
          <xsl:call-template name="tmpGetGroupRotXYVals">
            <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
            <xsl:with-param name="unit" select="$convertUnit"/>
            <xsl:with-param name="XY" select="'X'"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name="length" select="concat($tmp_X,$convertUnit)" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name ="y2">
        <xsl:variable name="tmp_Y">
          <xsl:call-template name="tmpGetGroupRotXYVals">
            <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
            <xsl:with-param name="unit" select="$convertUnit"/>
            <xsl:with-param name="XY" select="'Y'"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name="length" select="concat($tmp_Y,$convertUnit)" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="@draw:transform">
      
       
        <xsl:attribute name ="rot">
          <xsl:value-of select ="concat('draw-transform:ROT:',$width, ':',
																   $height, ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test="not(draw:enhanced-geometry/@draw:type='curvedLeftArrow' or
                    draw:enhanced-geometry/@draw:type='curvedRightArrow' or
                    draw:enhanced-geometry/@draw:type='curvedDownArrow' or
                    draw:enhanced-geometry/@draw:type='curvedUpArrow')">
        <xsl:if test="draw:enhanced-geometry/@draw:mirror-horizontal">
          <xsl:if test="draw:enhanced-geometry/@draw:mirror-horizontal='true'">
            <xsl:attribute name ="flipH">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:if>
      <xsl:if test="not(draw:enhanced-geometry/@draw:type='curvedLeftArrow' or
                          draw:enhanced-geometry/@draw:type='curvedRightArrow' or
                          draw:enhanced-geometry/@draw:type='curvedDownArrow' or
                          draw:enhanced-geometry/@draw:type='curvedUpArrow')">
        <xsl:if test="draw:enhanced-geometry/@draw:mirror-vertical">
          <xsl:if test="draw:enhanced-geometry/@draw:mirror-vertical='true'">
            <xsl:attribute name ="flipV">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:if>

      <!--Bug Fix for Shape Corner-Right Arrow from ODP to PPtx-->
      <xsl:if test="(draw:enhanced-geometry/@draw:enhanced-path='M 517 247 L 517 415 264 415 264 0 0 0 0 680 517 680 517 854 841 547 517 247 Z N')">
        <xsl:attribute name ="rot">
          <xsl:value-of select ="'5400000'"/>
        </xsl:attribute>
      </xsl:if>
      <!--End of bug fix code-->
      <a:off>
        <xsl:if test="@draw:transform">
          <xsl:attribute name ="x">
            <xsl:value-of select ="concat('draw-transform:X:',$width, ':',
																   $height, ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>

          <xsl:attribute name ="y">
            <xsl:value-of select ="concat('draw-transform:Y:',$width, ':',
																   $height, ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="not(@draw:transform)">
          <xsl:attribute name ="x">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="@svg:x"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name ="y">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="@svg:y"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </a:off>
      <a:ext>
        <xsl:attribute name ="cx">
               <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="@svg:width"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name ="cy">
            <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="@svg:height"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </a:ext>
    </xsl:template>
  <xsl:template name ="tmpGroupdrawCordinates">
    <xsl:variable name="convertUnit">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="@draw:transform"/>
    </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="width" >
      <xsl:choose>
        <xsl:when test="@svg:width =''">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:otherwise>
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:width"/>
      </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    
    </xsl:variable>
    <xsl:variable name="height" >
      <xsl:choose>
        <xsl:when test="@svg:height =''">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:otherwise>
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:height"/>
      </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
     
    </xsl:variable>
    <xsl:variable name="x" >
      <xsl:choose>
        <xsl:when test="@svg:x =''">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:otherwise>
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:x"/>
      </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
  
    </xsl:variable>
    <xsl:variable name="y" >
      <xsl:choose>
        <xsl:when test="@svg:y =''">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:otherwise>
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:y"/>
      </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
         <xsl:variable name ="angle">
       
        <xsl:variable name ="var_rot">
          <xsl:call-template name="tmpGetGroupRotXYVals">
            <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
            <xsl:with-param name="XY" select="'ROT'"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$var_rot=''">
          <xsl:value-of select="'0'"/>
        </xsl:if>
        <xsl:if test="$var_rot!=''">
          <xsl:value-of select="$var_rot"/>
        </xsl:if>
        </xsl:variable>
      <xsl:variable name ="x2">
        <xsl:variable name="tmp_X">
          <xsl:call-template name="tmpGetGroupRotXYVals">
            <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
            <xsl:with-param name="unit" select="$convertUnit"/>
            <xsl:with-param name="XY" select="'X'"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name="length" select="concat($tmp_X,$convertUnit)" />
        </xsl:call-template>
        <!--<xsl:if test="$tmp_X=''">
          <xsl:value-of select="'0'"/>
        </xsl:if>
        <xsl:if test="$tmp_X!=''">
          <xsl:value-of select="$tmp_X"/>
        </xsl:if>-->
      </xsl:variable>
    
      <xsl:variable name ="y2">
       
        <xsl:variable name="tmp_Y">
          <xsl:call-template name="tmpGetGroupRotXYVals">
            <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
            <xsl:with-param name="unit" select="$convertUnit"/>
            <xsl:with-param name="XY" select="'Y'"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="convertUnitsToCm">
          <xsl:with-param name="length" select="concat($tmp_Y,$convertUnit)" />
        </xsl:call-template>
        </xsl:variable>
      <xsl:if test="@draw:transform">
        <xsl:attribute name ="rot">
          <xsl:value-of select ="concat('draw-transform:ROT:',$width, ':',
																   $height, ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test="not(draw:enhanced-geometry/@draw:type='curvedLeftArrow' or
                    draw:enhanced-geometry/@draw:type='curvedRightArrow' or
                    draw:enhanced-geometry/@draw:type='curvedDownArrow' or
                    draw:enhanced-geometry/@draw:type='curvedUpArrow')">
        <xsl:if test="draw:enhanced-geometry/@draw:mirror-horizontal">
          <xsl:if test="draw:enhanced-geometry/@draw:mirror-horizontal='true'">
            <xsl:attribute name ="flipH">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:if>
      <xsl:if test="not(draw:enhanced-geometry/@draw:type='curvedLeftArrow' or
                          draw:enhanced-geometry/@draw:type='curvedRightArrow' or
                          draw:enhanced-geometry/@draw:type='curvedDownArrow' or
                          draw:enhanced-geometry/@draw:type='curvedUpArrow')">
        <xsl:if test="draw:enhanced-geometry/@draw:mirror-vertical">
          <xsl:if test="draw:enhanced-geometry/@draw:mirror-vertical='true'">
            <xsl:attribute name ="flipV">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:if>

      <!--Bug Fix for Shape Corner-Right Arrow from ODP to PPtx-->
      <xsl:if test="(draw:enhanced-geometry/@draw:enhanced-path='M 517 247 L 517 415 264 415 264 0 0 0 0 680 517 680 517 854 841 547 517 247 Z N')">
        <xsl:attribute name ="rot">
          <xsl:value-of select ="'5400000'"/>
        </xsl:attribute>
      </xsl:if>
      <!--End of bug fix code-->

        <a:off>
        <xsl:if test="@draw:transform">
          <xsl:attribute name ="x">
            <xsl:value-of select ="concat('draw-transform:XGroup:',$width, ':',
																   $height, ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>

          <xsl:attribute name ="y">
            <xsl:value-of select ="concat('draw-transform:YGroup:',$width, ':',
																   $height, ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="not(@draw:transform)">
          <xsl:attribute name ="x">
            <xsl:choose>
              <xsl:when test="$x=0 or contains($x,'NaN') or @svg:x='cm' or $x=''">
                <xsl:value-of select="'0'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select=" round(($x * 360000) div 1588)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name ="y">
            <xsl:choose>
              <xsl:when test="$y=0 or contains($y,'NaN') or @svg:y='cm' or $y=''">
                <xsl:value-of select="'0'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select=" round(($y * 360000) div 1588)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </a:off>
      <a:ext>
          <xsl:attribute name ="cx">
          <xsl:choose>
            <xsl:when test="$width='0' or contains($width,'NaN') or @svg:width='cm' or $width=''">
              <xsl:value-of select="'0'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="((number($width) * 360000) div 1588) &lt;= 1">
                  <xsl:value-of select="'0'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="format-number(( number($width) * 360000) div 1588,'#')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name ="cy">
          <xsl:choose>
            <xsl:when test="$height='0' or contains($height,'NaN') or @svg:height='cm' or $height=''">
              <xsl:value-of select="'0'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="((number($height) * 360000) div 1588) &lt;= 1">
                  <xsl:value-of select="'0'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="format-number(( number($height) * 360000) div 1588,'#')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
          </xsl:attribute>
      </a:ext>
     </xsl:template>
  <xsl:template name="tmpgetCustShapeType">
    <xsl:choose>
      <!-- Line -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path='M ?f0 ?f2 L ?f1 ?f3 N'">
        <xsl:value-of select ="'line'"/>
      </xsl:when>
      <!-- Text Box -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt202' and 
                      draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
        <xsl:value-of select ="'TextBox Custom '"/>
      </xsl:when>
      <!-- Oval -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='ellipse'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Oval')) ">
        <xsl:value-of select ="'Oval Custom '"/>
      </xsl:when>

      
      
      <!--#####################Arrows#############################-->
      <!-- U-Turn Arrow (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100' and draw:enhanced-geometry/@draw:enhanced-path='M 0 877824 L 0 384048 W 0 0 768096 768096 0 384048 384049 0 L 393192 0 W 9144 0 777240 768096 393192 0 777240 384049 L 777240 438912 L 886968 438912 L 667512 658368 L 448056 438912 L 557784 438912 L 557784 384048 A 228600 219456 557784 548640 557784 384048 393192 219456 L 384048 219456 A 219456 219456 548640 548640 384048 219456 219456 384048 L 219456 877824 Z N')
                        or draw:enhanced-geometry/@draw:enhanced-path = 'M ?f3 ?f6 L ?f3 ?f24 A ?f72 ?f73 ?f74 ?f75 ?f3 ?f24 ?f69 ?f71  W ?f76 ?f77 ?f78 ?f79 ?f3 ?f24 ?f69 ?f71 L ?f31 ?f5 A ?f115 ?f116 ?f117 ?f118 ?f31 ?f5 ?f112 ?f114  W ?f119 ?f120 ?f121 ?f122 ?f31 ?f5 ?f112 ?f114 L ?f22 ?f21 ?f4 ?f21 ?f28 ?f19 ?f29 ?f21 ?f30 ?f21 ?f30 ?f27 A ?f162 ?f163 ?f164 ?f165 ?f30 ?f27 ?f159 ?f161  W ?f166 ?f167 ?f168 ?f169 ?f30 ?f27 ?f159 ?f161 L ?f27 ?f15 A ?f198 ?f199 ?f200 ?f201 ?f27 ?f15 ?f195 ?f197  W ?f202 ?f203 ?f204 ?f205 ?f27 ?f15 ?f195 ?f197 L ?f15 ?f6 Z N'">
        <xsl:value-of select ="'U-Turn Arrow'"/>
      </xsl:when>
      <!-- Isosceles Triangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='isosceles-triangle' 
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f3 L ?f11 ?f2 ?f1 ?f3 Z N')">
        <xsl:value-of select ="'Isosceles Triangle '"/>
      </xsl:when>
      <!-- LeftUp Arrow -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt89' and
								    draw:enhanced-geometry/@draw:enhanced-path='M 0 ?f5 L ?f2 ?f0 ?f2 ?f7 ?f7 ?f7 ?f7 ?f2 ?f0 ?f2 ?f5 0 21600 ?f2 ?f1 ?f2 ?f1 ?f1 ?f2 ?f1 ?f2 21600 Z N')
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path ='M ?f0 ?f16 L ?f10 ?f13 ?f10 ?f20 ?f18 ?f20 ?f18 ?f10 ?f12 ?f10 ?f15 ?f2 ?f1 ?f10 ?f19 ?f10 ?f19 ?f21 ?f10 ?f21 ?f10 ?f3 Z N')">
        <xsl:value-of select ="'leftUpArrow'"/>
      </xsl:when>
      <!-- BentUp Arrow -->
 <!-- Sonata: 2710057 -PPTX: SP2 roundtrip arrows lost  -->
      <!-- Fix for the bug 24, Internal Defects.xls, date 9th Aug '07, by vijayeta-->
      <xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path='M 0 1428750 L 2562225 1428750 L 2562225 476250 L 2324100 476250 L 2800350 0 L 3276600 476250 L 3038475 476250 L 3038475 1905000 L 0 1905000 Z N'
                or draw:enhanced-geometry/@draw:enhanced-path='M 0 857250 L 790572 857250 L 790572 285750 L 647697 285750 L 933447 0 L 1219197 285750 L 1076322 285750 L 1076322 1143000 L 0 1143000 Z N'
                 or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f19 L ?f16 ?f19 ?f16 ?f10 ?f12 ?f10 ?f14 ?f2 ?f1 ?f10 ?f17 ?f10 ?f17 ?f3 ?f0 ?f3 Z N')">
        <xsl:value-of select ="'bentUpArrow'"/>
      </xsl:when>
      <!-- Quad Arrow -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='quad-arrow' 
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f6 L ?f14 ?f22 ?f14 ?f24 ?f19 ?f24 ?f19 ?f14 ?f16 ?f14 ?f9 ?f2 ?f17 ?f14 ?f20 ?f14 ?f20 ?f24 ?f21 ?f24 ?f21 ?f22 ?f1 ?f6 ?f21 ?f23 ?f21 ?f25 ?f20 ?f25 ?f20 ?f26 ?f17 ?f26 ?f9 ?f3 ?f16 ?f26 ?f19 ?f26 ?f19 ?f25 ?f14 ?f25 ?f14 ?f23 Z N')">
        <xsl:value-of select ="'Quad Arrow '"/>
      </xsl:when>
      <!-- Notched Right Arrow -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='notched-right-arrow'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f14 L ?f12 ?f14 ?f12 ?f2 ?f1 ?f6 ?f12 ?f3 ?f12 ?f15 ?f0 ?f15 ?f16 ?f6 Z N')	">
        <xsl:value-of select ="'notchedRightArrow'"/>
      </xsl:when>
      <xsl:when test ="draw:enhanced-geometry/@draw:type='curvedUpArrow'
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Curved Up Arrow'))">
        <xsl:value-of select ="'curvedUpArrow'"/>
      </xsl:when>
      <xsl:when test ="draw:enhanced-geometry/@draw:type='curvedRightArrow' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f3 ?f17 A ?f99 ?f100 ?f101 ?f102 ?f3 ?f17 ?f96 ?f98  W ?f103 ?f104 ?f105 ?f106 ?f3 ?f17 ?f96 ?f98 L ?f40 ?f36 ?f4 ?f39 ?f40 ?f37 ?f40 ?f33 A ?f146 ?f147 ?f148 ?f149 ?f40 ?f33 ?f143 ?f145  W ?f150 ?f151 ?f152 ?f153 ?f40 ?f33 ?f143 ?f145 Z N S M ?f4 ?f14 A ?f193 ?f194 ?f195 ?f196 ?f4 ?f14 ?f190 ?f192  W ?f197 ?f198 ?f199 ?f200 ?f4 ?f14 ?f190 ?f192 A ?f240 ?f241 ?f242 ?f243 ?f190 ?f192 ?f237 ?f239  W ?f244 ?f245 ?f246 ?f247 ?f190 ?f192 ?f237 ?f239 Z N F M ?f3 ?f17 A ?f99 ?f100 ?f101 ?f102 ?f3 ?f17 ?f96 ?f98  W ?f103 ?f104 ?f105 ?f106 ?f3 ?f17 ?f96 ?f98 L ?f40 ?f36 ?f4 ?f39 ?f40 ?f37 ?f40 ?f33 A ?f146 ?f147 ?f148 ?f149 ?f40 ?f33 ?f143 ?f145  W ?f150 ?f151 ?f152 ?f153 ?f40 ?f33 ?f143 ?f145 L ?f3 ?f17 A ?f268 ?f269 ?f270 ?f271 ?f3 ?f17 ?f265 ?f267  W ?f272 ?f273 ?f274 ?f275 ?f3 ?f17 ?f265 ?f267 L ?f4 ?f14 A ?f193 ?f194 ?f195 ?f196 ?f4 ?f14 ?f190 ?f192  W ?f197 ?f198 ?f199 ?f200 ?f4 ?f14 ?f190 ?f192 N')">
        <xsl:value-of select ="'curvedRightArrow'"/>
      </xsl:when>
      <xsl:when test ="draw:enhanced-geometry/@draw:type='curvedDownArrow' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f38 ?f6 L ?f35 ?f39 ?f31 ?f39 A ?f99 ?f100 ?f101 ?f102 ?f31 ?f39 ?f96 ?f98  W ?f103 ?f104 ?f105 ?f106 ?f31 ?f39 ?f96 ?f98 L ?f25 ?f5 A ?f146 ?f147 ?f148 ?f149 ?f25 ?f5 ?f143 ?f145  W ?f150 ?f151 ?f152 ?f153 ?f25 ?f5 ?f143 ?f145 L ?f36 ?f39 Z N S M ?f48 ?f47 A ?f193 ?f194 ?f195 ?f196 ?f48 ?f47 ?f190 ?f192  W ?f197 ?f198 ?f199 ?f200 ?f48 ?f47 ?f190 ?f192 L ?f3 ?f6 A ?f240 ?f241 ?f242 ?f243 ?f3 ?f6 ?f237 ?f239  W ?f244 ?f245 ?f246 ?f247 ?f3 ?f6 ?f237 ?f239 Z N F M ?f48 ?f47 A ?f193 ?f194 ?f195 ?f196 ?f48 ?f47 ?f190 ?f192  W ?f197 ?f198 ?f199 ?f200 ?f48 ?f47 ?f190 ?f192 L ?f3 ?f6 A ?f268 ?f269 ?f270 ?f271 ?f3 ?f6 ?f265 ?f267  W ?f272 ?f273 ?f274 ?f275 ?f3 ?f6 ?f265 ?f267 L ?f25 ?f5 A ?f146 ?f147 ?f148 ?f149 ?f25 ?f5 ?f143 ?f145  W ?f150 ?f151 ?f152 ?f153 ?f25 ?f5 ?f143 ?f145 L ?f36 ?f39 ?f38 ?f6 ?f35 ?f39 ?f31 ?f39 A ?f99 ?f100 ?f101 ?f102 ?f31 ?f39 ?f96 ?f98  W ?f103 ?f104 ?f105 ?f106 ?f31 ?f39 ?f96 ?f98 N')">
        <xsl:value-of select ="'curvedDownArrow'"/>
      </xsl:when>
      <xsl:when test ="draw:enhanced-geometry/@draw:type='curvedLeftArrow' 
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f3 ?f39 L ?f40 ?f36 ?f40 ?f32 A ?f97 ?f98 ?f99 ?f100 ?f40 ?f32 ?f94 ?f96  W ?f101 ?f102 ?f103 ?f104 ?f40 ?f32 ?f94 ?f96 A ?f144 ?f145 ?f146 ?f147 ?f94 ?f96 ?f141 ?f143  W ?f148 ?f149 ?f150 ?f151 ?f94 ?f96 ?f141 ?f143 L ?f40 ?f37 Z N S M ?f4 ?f26 A ?f191 ?f192 ?f193 ?f194 ?f4 ?f26 ?f188 ?f190  W ?f195 ?f196 ?f197 ?f198 ?f4 ?f26 ?f188 ?f190 L ?f3 ?f5 A ?f238 ?f239 ?f240 ?f241 ?f3 ?f5 ?f235 ?f237  W ?f242 ?f243 ?f244 ?f245 ?f3 ?f5 ?f235 ?f237 Z N F M ?f4 ?f26 A ?f191 ?f192 ?f193 ?f194 ?f4 ?f26 ?f188 ?f190  W ?f195 ?f196 ?f197 ?f198 ?f4 ?f26 ?f188 ?f190 L ?f3 ?f5 A ?f238 ?f239 ?f240 ?f241 ?f3 ?f5 ?f235 ?f237  W ?f242 ?f243 ?f244 ?f245 ?f3 ?f5 ?f235 ?f237 L ?f4 ?f26 A ?f266 ?f267 ?f268 ?f269 ?f4 ?f26 ?f263 ?f265  W ?f270 ?f271 ?f272 ?f273 ?f4 ?f26 ?f263 ?f265 L ?f40 ?f37 ?f3 ?f39 ?f40 ?f36 ?f40 ?f32 A ?f97 ?f98 ?f99 ?f100 ?f40 ?f32 ?f94 ?f96  W ?f101 ?f102 ?f103 ?f104 ?f40 ?f32 ?f94 ?f96 N')">
        <xsl:value-of select ="'curvedLeftArrow'"/>
      </xsl:when>
      <!-- Left-Right Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='left-right-arrow' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f6 L ?f13 ?f2 ?f13 ?f16 ?f14 ?f16 ?f14 ?f2 ?f1 ?f6 ?f14 ?f3 ?f14 ?f17 ?f13 ?f17 ?f13 ?f3 Z N')">
        <xsl:value-of select ="'Left-Right Arrow '"/>
      </xsl:when>
      <!-- Up-Down Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='up-down-arrow' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Up-Down Arrow'))">
        <xsl:value-of select ="'Up-Down Arrow '"/>
      </xsl:when>
      <!-- Right Arrow (Added by Mathi as on 4/7/2007 -->
      <xsl:when test = "draw:enhanced-geometry/@draw:type='right-arrow' 
                         or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 ?f0 L ?f1 ?f0 ?f1 0 21600 10800 ?f1 21600 ?f1 ?f2 0 ?f2 Z N') ">
        <xsl:value-of select ="'Right Arrow '"/>
      </xsl:when>
      <!-- Left Arrow (Added by Mathi as on 5/7/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='left-arrow' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 21600 ?f0 L ?f1 ?f0 ?f1 0 0 10800 ?f1 21600 ?f1 ?f2 21600 ?f2 Z N')">
        <xsl:value-of select ="'Left Arrow '"/>
      </xsl:when>
      <!-- Up Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='up-arrow' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 21600 L ?f0 ?f1 0 ?f1 10800 0 21600 ?f1 ?f2 ?f1 ?f2 21600 Z N')">
        <xsl:value-of select ="'Up Arrow '"/>
      </xsl:when>
      <!-- Down Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='down-arrow' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 0 L ?f0 ?f1 0 ?f1 10800 21600 21600 ?f1 ?f2 ?f1 ?f2 0 Z N')">
        <xsl:value-of select ="'Down Arrow '"/>
      </xsl:when>

      <!-- Circular Arrow -->
<!-- Sonata: 2710057 -PPTX: SP2 roundtrip arrows lost  -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='circular-arrow'
            or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 44649 357189 A ?f89 ?f90 ?f91 ?f92 44649 357189 ?f86 ?f88  W ?f93 ?f94 ?f95 ?f96 44649 357189 ?f86 ?f88 L 1012123 223630 982272 357191 833528 223630 857776 223630 A ?f136 ?f137 ?f138 ?f139 857776 223630 ?f133 ?f135  W ?f140 ?f141 ?f142 ?f143 857776 223630 ?f133 ?f135 Z N')
            or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 71438 800099 A ?f90 ?f91 ?f92 ?f93 71438 800099 ?f87 ?f89  W ?f94 ?f95 ?f96 ?f97 71438 800099 ?f87 ?f89 L 1132541 655800 1000125 800101 846791 655799 917680 655799 A ?f137 ?f138 ?f139 ?f140 917680 655799 ?f134 ?f136  W ?f141 ?f142 ?f143 ?f144 917680 655799 ?f134 ?f136 Z N')">
        <xsl:value-of select ="'circular-arrow'"/>
      </xsl:when>
      <!-- Bent Arrow (Added by A.Mathi as on 23/07/2007),2710057 -->
      <xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path='M 0 868680 L 0 457772 W 0 101727 712090 813817 0 457772 356046 101727 L 610362 101727 L 610362 0 L 813816 203454 L 610362 406908 L 610362 305181 L 356045 305181 A 203454 305181 508636 610363 356045 305181 203454 457772 L 203454 868680 Z N'
                 or (draw:enhanced-geometry/@draw:type='non-primitive' and                                                                     
                       draw:enhanced-geometry/@draw:enhanced-path ='M ?f3 ?f6 L ?f3 ?f27 A ?f67 ?f68 ?f69 ?f70 ?f3 ?f27 ?f64 ?f66  W ?f71 ?f72 ?f73 ?f74 ?f3 ?f27 ?f64 ?f66 L ?f19 ?f17 ?f19 ?f5 ?f4 ?f15 ?f19 ?f26 ?f19 ?f25 ?f24 ?f25 A ?f114 ?f115 ?f116 ?f117 ?f24 ?f25 ?f111 ?f113  W ?f118 ?f119 ?f120 ?f121 ?f24 ?f25 ?f111 ?f113 L ?f14 ?f6 Z N')">
        <xsl:value-of select ="'bentArrow '"/>
      </xsl:when>
      <!--Bug Fix for Shape Corner-Right Arrow from ODP to PPtx-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:enhanced-path='M 517 247 L 517 415 264 415 264 0 0 0 0 680 517 680 517 854 841 547 517 247 Z N')">
        <xsl:value-of select ="'bentUpArrow '"/>
      </xsl:when>






      <!--#####################         BASIC SHAPES     #############################-->

      <!-- Right Triangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='right-triangle' 
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f3 L ?f0 ?f2 ?f1 ?f3 Z N')">
        <xsl:value-of select ="'Right Triangle '"/>
      </xsl:when>
      <!-- Parallelogram -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='parallelogram'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f3 L ?f14 ?f2 ?f1 ?f2 ?f16 ?f3 Z N')">
        <xsl:value-of select ="'Parallelogram '"/>
      </xsl:when>
      <!-- Trapezoid (Added by A.Mathi as on 24/07/2007) -->
      <xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt100' and 
									draw:enhanced-geometry/@draw:enhanced-path='M 0 1216152 L 228600 0 L 685800 0 L 914400 1216152 Z N')
                   or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f3 L ?f14 ?f2 ?f15 ?f2 ?f1 ?f3 Z N')">
        <xsl:value-of select ="'Trapezoid '"/>
      </xsl:when>
      <xsl:when test="draw:enhanced-geometry/@draw:type='trapezoid'">
        <xsl:value-of select ="'flowchart-manual-operation '"/>
      </xsl:when>
      <!-- Diamond -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='diamond' and
									 draw:enhanced-geometry/@draw:enhanced-path='M 10800 0 L 21600 10800 10800 21600 0 10800 10800 0 Z N')
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f7 L ?f11 ?f2 ?f1 ?f7 ?f11 ?f3 Z N')	">
        <xsl:value-of select ="'Diamond '"/>
      </xsl:when>

      <!-- Cube -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='cube'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f0 ?f8 L ?f12 ?f8 ?f12 ?f3 ?f0 ?f3 Z N S M ?f12 ?f8 L ?f1 ?f2 ?f1 ?f9 ?f12 ?f3 Z N S M ?f0 ?f8 L ?f8 ?f2 ?f1 ?f2 ?f12 ?f8 Z N F M ?f0 ?f8 L ?f8 ?f2 ?f1 ?f2 ?f1 ?f9 ?f12 ?f3 ?f0 ?f3 Z M ?f0 ?f8 L ?f12 ?f8 ?f1 ?f2 M ?f12 ?f8 L ?f12 ?f3 N')	">
        <xsl:value-of select ="'Cube '"/>
      </xsl:when>
      <!-- Can -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='can'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f2 ?f13 A ?f55 ?f56 ?f57 ?f58 ?f2 ?f13 ?f52 ?f54  W ?f59 ?f60 ?f61 ?f62 ?f2 ?f13 ?f52 ?f54 L ?f3 ?f15 A ?f102 ?f103 ?f104 ?f105 ?f3 ?f15 ?f99 ?f101  W ?f106 ?f107 ?f108 ?f109 ?f3 ?f15 ?f99 ?f101 Z N S M ?f2 ?f13 A ?f126 ?f127 ?f128 ?f129 ?f2 ?f13 ?f123 ?f125  W ?f130 ?f131 ?f132 ?f133 ?f2 ?f13 ?f123 ?f125 A ?f142 ?f143 ?f144 ?f145 ?f123 ?f125 ?f140 ?f141  W ?f146 ?f147 ?f148 ?f149 ?f123 ?f125 ?f140 ?f141 Z N F M ?f3 ?f13 A ?f102 ?f154 ?f104 ?f155 ?f3 ?f13 ?f99 ?f153  W ?f106 ?f156 ?f108 ?f157 ?f3 ?f13 ?f99 ?f153 A ?f166 ?f167 ?f168 ?f169 ?f99 ?f153 ?f164 ?f165  W ?f170 ?f171 ?f172 ?f173 ?f99 ?f153 ?f164 ?f165 L ?f3 ?f15 A ?f102 ?f103 ?f104 ?f105 ?f3 ?f15 ?f99 ?f101  W ?f106 ?f107 ?f108 ?f109 ?f3 ?f15 ?f99 ?f101 L ?f2 ?f13 N')	">
        <xsl:value-of select ="'Can '"/>
      </xsl:when>


      <!-- Chord -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100' and 
                       draw:enhanced-geometry/@draw:enhanced-path='M 780489 780489 W 0 0 914400 914400 780489 780489 457200 0 Z N')
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f51 ?f52 A ?f122 ?f123 ?f124 ?f125 ?f51 ?f52 ?f119 ?f121  W ?f126 ?f127 ?f128 ?f129 ?f51 ?f52 ?f119 ?f121 Z N')	">
        <xsl:value-of select ="'Chord'"/>
      </xsl:when>

      <!--End of Fix for the bug 24, Internal Defects.xls, date 9th Aug '07, by vijayeta-->
      <!-- Cross (Added by Mathi on 19/7/2007)-->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='cross'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f8 L ?f8 ?f8 ?f8 ?f2 ?f9 ?f2 ?f9 ?f8 ?f1 ?f8 ?f1 ?f10 ?f9 ?f10 ?f9 ?f3 ?f8 ?f3 ?f8 ?f10 ?f0 ?f10 Z N')	">
        <xsl:value-of select ="'plus '"/>
      </xsl:when>
      <!-- "No Symbol" (Added by A.Mathi as on 19/07/2007)-->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='forbidden'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'No Smoking'))
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path  = 'M ?f3 ?f9 A ?f124 ?f125 ?f126 ?f127 ?f3 ?f9 ?f121 ?f123  W ?f128 ?f129 ?f130 ?f131 ?f3 ?f9 ?f121 ?f123 A ?f167 ?f168 ?f169 ?f170 ?f121 ?f123 ?f164 ?f166  W ?f171 ?f172 ?f173 ?f174 ?f121 ?f123 ?f164 ?f166 A ?f210 ?f211 ?f212 ?f213 ?f164 ?f166 ?f207 ?f209  W ?f214 ?f215 ?f216 ?f217 ?f164 ?f166 ?f207 ?f209 A ?f253 ?f254 ?f255 ?f256 ?f207 ?f209 ?f250 ?f252  W ?f257 ?f258 ?f259 ?f260 ?f207 ?f209 ?f250 ?f252 Z M ?f61 ?f62 A ?f291 ?f292 ?f293 ?f294 ?f61 ?f62 ?f288 ?f290  W ?f295 ?f296 ?f297 ?f298 ?f61 ?f62 ?f288 ?f290 Z M ?f63 ?f64 A ?f334 ?f335 ?f336 ?f337 ?f63 ?f64 ?f331 ?f333  W ?f338 ?f339 ?f340 ?f341 ?f63 ?f64 ?f331 ?f333 Z N')">
        <xsl:value-of select ="'noSmoking '"/>
      </xsl:when>     
		<xsl:when test ="(draw:enhanced-geometry/@draw:type='block-arc')
                     or (draw:enhanced-geometry/@draw:type='non-primitive' 
                          and (draw:enhanced-geometry/@draw:enhanced-path = 'M ?f55 ?f56 A ?f170 ?f171 ?f172 ?f173 ?f55 ?f56 ?f167 ?f169 W ?f174 ?f175 ?f176 ?f177 ?f55 ?f56 ?f167 ?f169 L ?f80 ?f81 A ?f208 ?f209 ?f210 ?f211 ?f80 ?f81 ?f205 ?f207 W ?f212 ?f213 ?f214 ?f215 ?f80 ?f81 ?f205 ?f207 Z N' 
                               or  draw:enhanced-geometry/@draw:enhanced-path =  'M ?f55 ?f56 A ?f170 ?f171 ?f172 ?f173 ?f55 ?f56 ?f167 ?f169  W ?f174 ?f175 ?f176 ?f177 ?f55 ?f56 ?f167 ?f169 L ?f80 ?f81 A ?f208 ?f209 ?f210 ?f211 ?f80 ?f81 ?f205 ?f207  W ?f212 ?f213 ?f214 ?f215 ?f80 ?f81 ?f205 ?f207 Z N'))	">
			<xsl:value-of select ="'Block Arc '"/>
		</xsl:when>		
		<!-- Pentagon -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='pentagon-right'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f34 ?f38 L ?f11 ?f4 ?f37 ?f38 ?f36 ?f39 ?f35 ?f39 Z N')	">
                                                                                                                             
        <xsl:value-of select ="'Pentagon '"/>
		</xsl:when>
      <!-- homePlate -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f2 L ?f11 ?f2 ?f1 ?f6 ?f11 ?f3 ?f0 ?f3 Z N'	">

			<xsl:value-of select ="'Pentagon '"/>
		</xsl:when>
		<!-- Chevron -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='chevron'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Chevron'))	">
			<xsl:value-of select ="'Chevron '"/>
		</xsl:when>
		
      <!-- Heptagon -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='HEPTAGON'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f19 ?f26 L ?f20 ?f25 ?f9 ?f2 ?f23 ?f25 ?f24 ?f26 ?f22 ?f27 ?f21 ?f27 Z N')">
        <xsl:value-of select="'Heptagon '"/>
      </xsl:when>
      <!-- Regular Pentagon -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='pentagon' and
									 draw:enhanced-geometry/@draw:enhanced-path='M 10800 0 L 0 8260 4230 21600 17370 21600 21600 8260 10800 0 Z N')
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f34 ?f38 L ?f11 ?f4 ?f37 ?f38 ?f36 ?f39 ?f35 ?f39 Z N')">
        <xsl:value-of select ="'Regular Pentagon '"/>
      </xsl:when>
      <!-- Hexagon -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='hexagon'
                         or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f2 ?f8 L ?f15 ?f23 ?f16 ?f23 ?f3 ?f8 ?f16 ?f24 ?f15 ?f24 Z N')">
        <xsl:value-of select ="'Hexagon '"/>
      </xsl:when>
      <!-- Octagon -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='octagon'
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f8 L ?f8 ?f2 ?f9 ?f2 ?f1 ?f8 ?f1 ?f10 ?f9 ?f3 ?f8 ?f3 ?f0 ?f10 Z N')">
        <xsl:value-of select ="'Octagon '"/>
      </xsl:when>
      <!-- Decagon -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='DECAGON'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Decagon'))">
        <xsl:value-of select="'Decagon '"/>
      </xsl:when>
      <!-- Dodecagon -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='DODECAGON'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Dodecagon'))">
        <xsl:value-of select="'Dodecagon '"/>
      </xsl:when>
      <!-- Half Frame -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='HALFFRAME'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f2 L ?f1 ?f2 ?f12 ?f10 ?f8 ?f10 ?f8 ?f14 ?f0 ?f3 Z N')">
        <xsl:value-of select="'Half Frame '"/>
      </xsl:when>
      <!-- Pie -->
      <xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt100' and 
						draw:enhanced-geometry/@draw:enhanced-path='V 0 0 21600 21600 ?f5 ?f7 ?f1 ?f3 L 10800 10800 Z N')
             or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f35 ?f36 A ?f120 ?f121 ?f122 ?f123 ?f35 ?f36 ?f117 ?f119  W ?f124 ?f125 ?f126 ?f127 ?f35 ?f36 ?f117 ?f119 L ?f11 ?f8 Z N')">
        <xsl:value-of select="'Pie '"/>
      </xsl:when>
      <!-- Frame -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='frame'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z M ?f8 ?f8 L ?f8 ?f10 ?f9 ?f10 ?f9 ?f8 Z N')">
        <xsl:value-of select="'Frame '"/>
      </xsl:when>
      <!-- L-Shape -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='CORNER'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f2 L ?f9 ?f2 ?f9 ?f11 ?f1 ?f11 ?f1 ?f3 ?f0 ?f3 Z N')">
        <xsl:value-of select="'L-Shape '"/>
      </xsl:when>
      <!-- Diagonal Stripe -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='diagstripe'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f14 L ?f11 ?f2 ?f1 ?f2 ?f0 ?f3 Z N')">
        <xsl:value-of select="'Diagonal Stripe '"/>
      </xsl:when>
      <!-- Plaque -->
      <xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt21' and 
						draw:enhanced-geometry/@draw:enhanced-path='M ?f0 0 Y 0 ?f1 L 0 ?f2 X ?f0 21600 L ?f3 21600 Y 21600 ?f2 L 21600 ?f1 X ?f3 0 Z N')
            or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Plaque'))">
        <xsl:value-of select="'Plaque '"/>
      </xsl:when>
      <!-- Bevel -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='quad-bevel'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f12 ?f12 L ?f13 ?f12 ?f13 ?f14 ?f12 ?f14 Z N S M ?f0 ?f2 L ?f1 ?f2 ?f13 ?f12 ?f12 ?f12 Z N S M ?f0 ?f3 L ?f12 ?f14 ?f13 ?f14 ?f1 ?f3 Z N S M ?f0 ?f2 L ?f12 ?f12 ?f12 ?f14 ?f0 ?f3 Z N S M ?f1 ?f2 L ?f1 ?f3 ?f13 ?f14 ?f13 ?f12 Z N F M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z M ?f12 ?f12 L ?f13 ?f12 ?f13 ?f14 ?f12 ?f14 Z M ?f0 ?f2 L ?f12 ?f12 M ?f0 ?f3 L ?f12 ?f14 M ?f1 ?f2 L ?f13 ?f12 M ?f1 ?f3 L ?f13 ?f14 N')">
        <xsl:value-of select="'Bevel '"/>
      </xsl:when>
      <!-- Donut -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='ring'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f3 ?f9 A ?f79 ?f80 ?f81 ?f82 ?f3 ?f9 ?f76 ?f78  W ?f83 ?f84 ?f85 ?f86 ?f3 ?f9 ?f76 ?f78 A ?f122 ?f123 ?f124 ?f125 ?f76 ?f78 ?f119 ?f121  W ?f126 ?f127 ?f128 ?f129 ?f76 ?f78 ?f119 ?f121 A ?f165 ?f166 ?f167 ?f168 ?f119 ?f121 ?f162 ?f164  W ?f169 ?f170 ?f171 ?f172 ?f119 ?f121 ?f162 ?f164 A ?f208 ?f209 ?f210 ?f211 ?f162 ?f164 ?f205 ?f207  W ?f212 ?f213 ?f214 ?f215 ?f162 ?f164 ?f205 ?f207 Z M ?f17 ?f9 A ?f248 ?f249 ?f250 ?f251 ?f17 ?f9 ?f245 ?f247  W ?f252 ?f253 ?f254 ?f255 ?f17 ?f9 ?f245 ?f247 A ?f284 ?f285 ?f286 ?f287 ?f245 ?f247 ?f281 ?f283  W ?f288 ?f289 ?f290 ?f291 ?f245 ?f247 ?f281 ?f283 A ?f320 ?f321 ?f322 ?f323 ?f281 ?f283 ?f317 ?f319  W ?f324 ?f325 ?f326 ?f327 ?f281 ?f283 ?f317 ?f319 A ?f356 ?f357 ?f358 ?f359 ?f317 ?f319 ?f353 ?f355  W ?f360 ?f361 ?f362 ?f363 ?f317 ?f319 ?f353 ?f355 Z N')">
        <xsl:value-of select="'Donut '"/>
      </xsl:when>
      <!-- TearDrop -->
      <xsl:when test="draw:enhanced-geometry/@draw:type='TEARDROP'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Teardrop')) ">
        <xsl:value-of select="'Teardrop '"/>
      </xsl:when>
      
     
      <!--End of bug fix code-->
      
      <!--  Folded Corner (Added by A.Mathi as on 19/07/2007),2710057 -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='paper'
                   or (draw:enhanced-geometry/@draw:type='non-primitive' and  draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f12 ?f10 ?f3 ?f0 ?f3 Z N S M ?f10 ?f3 L ?f11 ?f13 ?f1 ?f12 Z N F M ?f10 ?f3 L ?f11 ?f13 ?f1 ?f12 ?f10 ?f3 ?f0 ?f3 ?f0 ?f2 ?f1 ?f2 ?f1 ?f12 N')">
        <xsl:value-of select ="'foldedCorner '"/>
      </xsl:when>
		<!--  Lightning Bolt (Added by A.Mathi as on 20/07/2007) -->
		<xsl:when test ="(draw:enhanced-geometry/@draw:enhanced-path='M 640 233 L 221 293 506 12 367 0 29 406 431 347 145 645 99 520 0 861 326 765 209 711 640 233 640 233 Z N' or 
					     draw:enhanced-geometry/@draw:enhanced-path='M 8458 0 L 0 3923 7564 8416 4993 9720 12197 13904 9987 14934 21600 21600 14768 12911 16558 12016 11030 6840 12831 6120 8458 0 Z N')
               or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 8472 0 L 12860 6080 11050 6797 16577 12007 14767 12877 21600 21600 10012 14915 12222 13987 5022 9705 7602 8382 0 3890 Z N') ">
			<xsl:value-of select ="'lightningBolt '"/>
		</xsl:when>
		<!--  Explosion 1 (Modified by A.Mathi) -->
		<xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt71' and
					   draw:enhanced-geometry/@draw:enhanced-path='M 10901 5905 L 8458 2399 7417 6425 476 2399 4732 7722 106 8718 3828 11880 243 14689 5772 14041 4868 17719 7819 15730 8590 21600 10637 15038 13349 19840 14125 14561 18248 18195 16938 13044 21600 13393 17710 10579 21198 8242 16806 7417 18482 4560 14257 5429 14623 106 10901 5905 Z N')
            or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 10800 5800 L 14522 0 14155 5325 18380 4457 16702 7315 21097 8137 17607 10475 21600 13290 16837 12942 18145 18095 14020 14457 13247 19737 10532 14935 8485 21600 7715 15627 4762 17617 5667 13937 135 14587 3722 11775 0 8615 4627 7617 370 2295 7312 6320 8352 2295 Z N') ">
			<xsl:value-of select ="'irregularSeal1 '"/>
		</xsl:when>
      <!-- Left Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='left-bracket'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Left Bracket')) ">
        <xsl:value-of select ="'Left Bracket '"/>
      </xsl:when>
      <!-- Right Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='right-bracket'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Right Bracket')) ">
        <xsl:value-of select ="'Right Bracket '"/>
      </xsl:when>
      <!-- Left Brace (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='left-brace'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f3 ?f5 A ?f68 ?f69 ?f70 ?f71 ?f3 ?f5 ?f65 ?f67  W ?f72 ?f73 ?f74 ?f75 ?f3 ?f5 ?f65 ?f67 L ?f9 ?f16 A ?f115 ?f116 ?f117 ?f118 ?f9 ?f16 ?f112 ?f114  W ?f119 ?f120 ?f121 ?f122 ?f9 ?f16 ?f112 ?f114 A ?f145 ?f146 ?f147 ?f148 ?f112 ?f114 ?f142 ?f144  W ?f149 ?f150 ?f151 ?f152 ?f112 ?f114 ?f142 ?f144 L ?f9 ?f14 A ?f188 ?f189 ?f190 ?f191 ?f9 ?f14 ?f185 ?f187  W ?f192 ?f193 ?f194 ?f195 ?f9 ?f14 ?f185 ?f187 Z N F M ?f3 ?f5 A ?f68 ?f69 ?f70 ?f71 ?f3 ?f5 ?f65 ?f67  W ?f72 ?f73 ?f74 ?f75 ?f3 ?f5 ?f65 ?f67 L ?f9 ?f16 A ?f115 ?f116 ?f117 ?f118 ?f9 ?f16 ?f112 ?f114  W ?f119 ?f120 ?f121 ?f122 ?f9 ?f16 ?f112 ?f114 A ?f145 ?f146 ?f147 ?f148 ?f112 ?f114 ?f142 ?f144  W ?f149 ?f150 ?f151 ?f152 ?f112 ?f114 ?f142 ?f144 L ?f9 ?f14 A ?f188 ?f189 ?f190 ?f191 ?f9 ?f14 ?f185 ?f187  W ?f192 ?f193 ?f194 ?f195 ?f9 ?f14 ?f185 ?f187 N') ">
        <xsl:value-of select ="'Left Brace '"/>
      </xsl:when>
      <!-- Right Brace (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='right-brace'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f3 ?f5 A ?f70 ?f71 ?f72 ?f73 ?f3 ?f5 ?f67 ?f69  W ?f74 ?f75 ?f76 ?f77 ?f3 ?f5 ?f67 ?f69 L ?f10 ?f17 A ?f117 ?f118 ?f119 ?f120 ?f10 ?f17 ?f114 ?f116  W ?f121 ?f122 ?f123 ?f124 ?f10 ?f17 ?f114 ?f116 A ?f147 ?f148 ?f149 ?f150 ?f114 ?f116 ?f144 ?f146  W ?f151 ?f152 ?f153 ?f154 ?f114 ?f116 ?f144 ?f146 L ?f10 ?f18 A ?f190 ?f191 ?f192 ?f193 ?f10 ?f18 ?f187 ?f189  W ?f194 ?f195 ?f196 ?f197 ?f10 ?f18 ?f187 ?f189 Z N F M ?f3 ?f5 A ?f70 ?f71 ?f72 ?f73 ?f3 ?f5 ?f67 ?f69  W ?f74 ?f75 ?f76 ?f77 ?f3 ?f5 ?f67 ?f69 L ?f10 ?f17 A ?f117 ?f118 ?f119 ?f120 ?f10 ?f17 ?f114 ?f116  W ?f121 ?f122 ?f123 ?f124 ?f10 ?f17 ?f114 ?f116 A ?f147 ?f148 ?f149 ?f150 ?f114 ?f116 ?f144 ?f146  W ?f151 ?f152 ?f153 ?f154 ?f114 ?f116 ?f144 ?f146 L ?f10 ?f18 A ?f190 ?f191 ?f192 ?f193 ?f10 ?f18 ?f187 ?f189  W ?f194 ?f195 ?f196 ?f197 ?f10 ?f18 ?f187 ?f189 N') ">
        <xsl:value-of select ="'Right Brace '"/>
      </xsl:when>
      <!-- Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='rectangular-callout'
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 0 L 0 3590 ?f2 ?f3 0 8970 0 12630 ?f4 ?f5 0 18010 0 21600 3590 21600 ?f6 ?f7 8970 21600 12630 21600 ?f8 ?f9 18010 21600 21600 21600 21600 18010 ?f10 ?f11 21600 12630 21600 8970 ?f12 ?f13 21600 3590 21600 0 18010 0 ?f14 ?f15 12630 0 8970 0 ?f16 ?f17 3590 0 0 0 Z N') ">
        <xsl:value-of select ="'Rectangular Callout '"/>
      </xsl:when>
      <!-- Rounded Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='round-rectangular-callout'
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 3590 0 X 0 3590 L ?f2 ?f3 0 8970 0 12630 ?f4 ?f5 0 18010 Y 3590 21600 L ?f6 ?f7 8970 21600 12630 21600 ?f8 ?f9 18010 21600 X 21600 18010 L ?f10 ?f11 21600 12630 21600 8970 ?f12 ?f13 21600 3590 Y 18010 0 L ?f14 ?f15 12630 0 8970 0 ?f16 ?f17 Z N') ">
        <xsl:value-of select ="'wedgeRoundRectCallout '"/>
      </xsl:when>
      <!-- Oval Callout (Added by A.Mathi as on 23/07/2007) -->
	  <xsl:when test ="(draw:enhanced-geometry/@draw:type='round-callout' and draw:enhanced-geometry[@draw:enhanced-path != 'V 0 0 21600 21600 ?f2 ?f3 ?f2 ?f4 N'])
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'W 0 0 21600 21600 ?f22 ?f23 ?f18 ?f19 L ?f14 ?f15 Z N')">
		  <xsl:value-of select ="'wedgeEllipseCallout '"/>
	  </xsl:when>
      <!-- Cloud Callout (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='cloud-callout'
                        or (draw:enhanced-geometry/@draw:type='non-primitive' 
                            and draw:enhanced-geometry/@draw:enhanced-path = 'M 1930 7160 C 1530 4490 3400 1970 5270 1970 5860 1950 6470 2210 6970 2600 7450 1390 8340 650 9340 650 10004 690 10710 1050 11210 1700 11570 630 12330 0 13150 0 13840 0 14470 460 14870 1160 15330 440 16020 0 16740 0 17910 0 18900 1130 19110 2710 20240 3150 21060 4580 21060 6220 21060 6720 21000 7200 20830 7660 21310 8460 21600 9450 21600 10460 21600 12750 20310 14680 18650 15010 18650 17200 17370 18920 15770 18920 15220 18920 14700 18710 14240 18310 13820 20240 12490 21600 11000 21600 9890 21600 8840 20790 8210 19510 7620 20000 7930 20290 6240 20290 4850 20290 3570 19280 2900 17640 1300 17600 480 16300 480 14660 480 13900 690 13210 1070 12640 380 12160 0 11210 0 10120 0 8590 840 7330 1930 7160 Z N M 1930 7160 C 1950 7410 2040 7690 2090 7920 F N M 6970 2600 C 7200 2790 7480 3050 7670 3310 F N M 11210 1700 C 11130 1910 11080 2160 11030 2400 F N M 14870 1160 C 14720 1400 14640 1720 14540 2010 F N M 19110 2710 C 19130 2890 19230 3290 19190 3380 F N M 20830 7660 C 20660 8170 20430 8620 20110 8990 F N M 18660 15010 C 18740 14200 18280 12200 17000 11450 F N M 14240 18310 C 14320 17980 14350 17680 14370 17360 F N M 8220 19510 C 8060 19250 7960 18950 7860 18640 F N M 2900 17640 C 3090 17600 3280 17540 3460 17450 F N M 1070 12640 C 1400 12900 1780 13130 2330 13040 F N U ?f17 ?f18 1800 1800 0 23592960 Z N U ?f19 ?f20 1200 1200 0 23592960 Z N U ?f13 ?f14 700 700 0 23592960 Z N')">
        <xsl:value-of select ="'Cloud Callout '"/>
      </xsl:when>
		<!-- Line Callout 1) -->
		<xsl:when test ="(draw:enhanced-geometry/@draw:type='line-callout-1' or draw:enhanced-geometry/@draw:type='line-callout-3')
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z N F M ?f9 ?f2 Z L ?f9 ?f3 N F M ?f9 ?f8 L ?f11 ?f10 ?f13 ?f12 ?f15 ?f14 N') ">
			<xsl:value-of select ="'borderCallout1'"/>
		</xsl:when>
		<!--Line Callout 2)-->
		<xsl:when test ="draw:enhanced-geometry/@draw:type='line-callout-2'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z N F M ?f9 ?f8 L ?f11 ?f10 ?f13 ?f12 N') ">
			<xsl:value-of select ="'borderCallout2'"/>
		</xsl:when>
		<!--Line Callout 3)-->
		<xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt49') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M ?f6 ?f7 F L ?f4 ?f5 ?f2 ?f3 ?f0 ?f1 N')">
			<xsl:value-of select ="'borderCallout3'"/>
		</xsl:when>	

      
      
      
      <!--#####################         RECTANGLE     #############################-->

      <!-- Rounded Rectangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='round-rectangle' 
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N') ">
        <xsl:value-of select ="'Rounded Rectangle '"/>
      </xsl:when>
      <!-- Snip Single Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-card' and 
									    draw:enhanced-geometry/@draw:mirror-horizontal='true' and
									 draw:enhanced-geometry/@draw:enhanced-path='M 4300 0 L 21600 0 21600 21600 0 21600 0 4300 4300 0 Z N') 
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f0 ?f2 L ?f9 ?f2 ?f1 ?f8 ?f1 ?f3 ?f0 ?f3 Z N') ">
        <xsl:value-of select ="'Snip Single Corner Rectangle '"/>
      </xsl:when>
		<!-- Snip Same Side Corner Rectangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='snip2samerect'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f9 ?f2 L ?f10 ?f2 ?f1 ?f9 ?f1 ?f13 ?f12 ?f3 ?f11 ?f3 ?f0 ?f13 ?f0 ?f9 Z N') ">
			<xsl:value-of select ="'Snip Same Side Corner Rectangle '"/>
		</xsl:when>
		<!-- Snip Diagonal Corner Rectangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='snip2diagrect'
                         or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f9 ?f2 L ?f13 ?f2 ?f1 ?f12 ?f1 ?f11 ?f10 ?f3 ?f12 ?f3 ?f0 ?f14 ?f0 ?f9 Z N') ">
			<xsl:value-of select ="'Snip Diagonal Corner Rectangle '"/>
		</xsl:when>
		<!-- Snip and Round Single Corner Rectangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='sniproundrect'
                         or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f12 ?f4 L ?f14 ?f4 ?f3 ?f13 ?f3 ?f5 ?f2 ?f5 ?f2 ?f12 A ?f56 ?f57 ?f58 ?f59 ?f2 ?f12 ?f53 ?f55  W ?f60 ?f61 ?f62 ?f63 ?f2 ?f12 ?f53 ?f55 Z N') ">
			<xsl:value-of select ="'Snip and Round Single Corner Rectangle '"/>
		</xsl:when>

		<!-- Round Single Corner Rectangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='round1rect'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f3 ?f5 L ?f13 ?f5 A ?f55 ?f56 ?f57 ?f58 ?f13 ?f5 ?f52 ?f54  W ?f59 ?f60 ?f61 ?f62 ?f13 ?f5 ?f52 ?f54 L ?f4 ?f6 ?f3 ?f6 Z N') ">
			<xsl:value-of select ="'Round Single Corner Rectangle '"/>
		</xsl:when>
		<!-- Round Same Side Corner Rectangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='round2samerect'
                          or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f13 ?f5 L ?f14 ?f5 A ?f62 ?f63 ?f64 ?f65 ?f14 ?f5 ?f59 ?f61  W ?f66 ?f67 ?f68 ?f69 ?f14 ?f5 ?f59 ?f61 L ?f4 ?f16 A ?f105 ?f106 ?f107 ?f108 ?f4 ?f16 ?f102 ?f104  W ?f109 ?f110 ?f111 ?f112 ?f4 ?f16 ?f102 ?f104 L ?f15 ?f6 A ?f148 ?f149 ?f150 ?f151 ?f15 ?f6 ?f145 ?f147  W ?f152 ?f153 ?f154 ?f155 ?f15 ?f6 ?f145 ?f147 L ?f3 ?f13 A ?f191 ?f192 ?f193 ?f194 ?f3 ?f13 ?f188 ?f190  W ?f195 ?f196 ?f197 ?f198 ?f3 ?f13 ?f188 ?f190 Z N') ">
			<xsl:value-of select ="'Round Same Side Corner Rectangle '"/>
		</xsl:when>
		<!-- Round Diagonal Corner Rectangle -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='round2diagrect'
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f13 ?f5 L ?f16 ?f5 A ?f62 ?f63 ?f64 ?f65 ?f16 ?f5 ?f59 ?f61  W ?f66 ?f67 ?f68 ?f69 ?f16 ?f5 ?f59 ?f61 L ?f4 ?f14 A ?f105 ?f106 ?f107 ?f108 ?f4 ?f14 ?f102 ?f104  W ?f109 ?f110 ?f111 ?f112 ?f4 ?f14 ?f102 ?f104 L ?f15 ?f6 A ?f148 ?f149 ?f150 ?f151 ?f15 ?f6 ?f145 ?f147  W ?f152 ?f153 ?f154 ?f155 ?f15 ?f6 ?f145 ?f147 L ?f3 ?f13 A ?f191 ?f192 ?f193 ?f194 ?f3 ?f13 ?f188 ?f190  W ?f195 ?f196 ?f197 ?f198 ?f3 ?f13 ?f188 ?f190 Z N') ">
			<xsl:value-of select ="'Round Diagonal Corner Rectangle '"/>
		</xsl:when>

      <!-- Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='rectangle' and draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 0 L 21600 0 21600 21600 0 21600 Z N') ">
        <xsl:value-of select ="'Rectangle Custom '"/>
		</xsl:when>
      
      
      
      
      
		<!--###############   Equation Shapes   ####################-->
		<!--Not Equal-->
		<xsl:when test="draw:enhanced-geometry/@draw:type='mathnotequal'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f22 ?f26 L ?f39 ?f26 ?f55 ?f59 ?f54 ?f58 ?f47 ?f26 ?f23 ?f26 ?f23 ?f24 ?f48 ?f24 ?f49 ?f25 ?f23 ?f25 ?f23 ?f27 ?f50 ?f27 ?f61 ?f63 ?f60 ?f62 ?f45 ?f27 ?f22 ?f27 ?f22 ?f25 ?f43 ?f25 ?f41 ?f24 ?f22 ?f24 Z N')">
			<xsl:value-of select="'Not Equal '"/>
		</xsl:when>
		<!--Equal-->
		<xsl:when test="draw:enhanced-geometry/@draw:type='mathequal'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f19 ?f17 L ?f20 ?f17 ?f20 ?f15 ?f19 ?f15 Z M ?f19 ?f16 L ?f20 ?f16 ?f20 ?f18 ?f19 ?f18 Z N')">
			<xsl:value-of select="'Equal '"/>
		</xsl:when>
		<!--Plus-->
      <xsl:when test="draw:enhanced-geometry/@draw:type='mathplus'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f15 ?f20 L ?f16 ?f20 ?f16 ?f19 ?f17 ?f19 ?f17 ?f20 ?f18 ?f20 ?f18 ?f21 ?f17 ?f21 ?f17 ?f22 ?f16 ?f22 ?f16 ?f21 ?f15 ?f21 Z N')">
        <xsl:value-of select="'Plus '"/>
		</xsl:when>
		<!--Minus-->
		<xsl:when test="draw:enhanced-geometry/@draw:type='mathminus'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f15 ?f13 L ?f16 ?f13 ?f16 ?f14 ?f15 ?f14 Z N')">
			<xsl:value-of select="'Minus '"/>
		</xsl:when>
		<!--Multiply-->
		<xsl:when test="draw:enhanced-geometry/@draw:type='mathmultiply'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f40 ?f41 L ?f42 ?f43 ?f11 ?f46 ?f47 ?f43 ?f48 ?f41 ?f51 ?f8 ?f48 ?f53 ?f47 ?f54 ?f11 ?f55 ?f42 ?f54 ?f40 ?f53 ?f52 ?f8 Z N')">
			<xsl:value-of select="'Multiply '"/>
		</xsl:when>
		<!--Division-->
		<xsl:when test="draw:enhanced-geometry/@draw:type='mathdivide'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f12 ?f25 A ?f68 ?f69 ?f70 ?f71 ?f12 ?f25 ?f65 ?f67  W ?f72 ?f73 ?f74 ?f75 ?f12 ?f25 ?f65 ?f67 Z M ?f12 ?f26 A ?f111 ?f112 ?f113 ?f114 ?f12 ?f26 ?f108 ?f110  W ?f115 ?f116 ?f117 ?f118 ?f12 ?f26 ?f108 ?f110 Z M ?f27 ?f21 L ?f28 ?f21 ?f28 ?f22 ?f27 ?f22 Z N')">
			<xsl:value-of select="'Division '"/>
		</xsl:when>

		
		<!-- Explosion 2 -->
		<xsl:when test="draw:enhanced-geometry/@draw:type='bang'
                    or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 11462 4342 L 14790 0 14525 5777 18007 3172 16380 6532 21600 6645 16985 9402 18270 11290 16380 12310 18877 15632 14640 14350 14942 17370 12180 15935 11612 18842 9872 17370 8700 19712 7527 18125 4917 21600 4805 18240 1285 17825 3330 15370 0 12877 3935 11592 1172 8270 5372 7817 4502 3625 8550 6382 9722 1887 Z N')">
			<xsl:value-of select="'Explosion 2 '"/>
		</xsl:when>


      <!--#####################         FLOW CHART SHAPES     #############################-->
      <!-- Flow Chart: Process -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-process'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 0 L 1 0 1 1 0 1 Z N')">
        <xsl:value-of select ="'Flowchart: Process '"/>
      </xsl:when>
      <!-- Flow Chart: Alternate Process -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-alternate-process'
                         or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f3 ?f10 A ?f56 ?f57 ?f58 ?f59 ?f3 ?f10 ?f53 ?f55  W ?f60 ?f61 ?f62 ?f63 ?f3 ?f10 ?f53 ?f55 L ?f12 ?f5 A ?f99 ?f100 ?f101 ?f102 ?f12 ?f5 ?f96 ?f98  W ?f103 ?f104 ?f105 ?f106 ?f12 ?f5 ?f96 ?f98 L ?f4 ?f13 A ?f142 ?f143 ?f144 ?f145 ?f4 ?f13 ?f139 ?f141  W ?f146 ?f147 ?f148 ?f149 ?f4 ?f13 ?f139 ?f141 L ?f10 ?f6 A ?f185 ?f186 ?f187 ?f188 ?f10 ?f6 ?f182 ?f184  W ?f189 ?f190 ?f191 ?f192 ?f10 ?f6 ?f182 ?f184 Z N')">
        <xsl:value-of select ="'Flowchart: Alternate Process '"/>
      </xsl:when>
      <!-- Flow Chart: Decision -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-decision'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 1 L 1 0 2 1 1 2 Z N')">
        <xsl:value-of select ="'Flowchart: Decision '"/>
      </xsl:when>
      <!-- Flow Chart: Data -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-data'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 5 L 1 0 5 0 4 5 Z N')">
        <xsl:value-of select ="'Flowchart: Data '"/>
      </xsl:when>
      <!-- Flow Chart: Predefined-process -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-predefined-process'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M 0 0 L 1143000 0 1143000 990600 0 990600 Z N F M 142875 0 L 142875 990600 M 1000125 0 L 1000125 990600 N F M 0 0 L 1143000 0 1143000 990600 0 990600 Z N')
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path ='S M 0 0 L 2357454 0 2357454 785818 0 785818 Z N F M 294682 0 L 294682 785818 M 2062772 0 L 2062772 785818 N F M 0 0 L 2357454 0 2357454 785818 0 785818 Z N')">
        <xsl:value-of select ="'Flowchart: Predefined Process '"/>
      </xsl:when>
      <!-- Flow Chart: Internal-storage -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-internal-storage'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M 0 0 L 1219200 0 1219200 990600 0 990600 Z N F M 152400 0 L 152400 990600 M 0 123825 L 1219200 123825 N F M 0 0 L 1219200 0 1219200 990600 0 990600 Z N')
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M 0 0 L 1357322 0 1357322 1214446 0 1214446 Z N F M 169665 0 L 169665 1214446 M 0 151806 L 1357322 151806 N F M 0 0 L 1357322 0 1357322 1214446 0 1214446 Z N') ">
        <xsl:value-of select ="'Flowchart: Internal Storage '"/>
      </xsl:when>
      <!-- Flow Chart: Document -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-document'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 0 L 21600 0 21600 17322 C 10800 17322 10800 23922 0 20172 Z N')">
        <xsl:value-of select ="'Flowchart: Document '"/>
      </xsl:when>
      <!-- Flow Chart: Multidocument -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-multidocument'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M 0 20782 C 9298 23542 9298 18022 18595 18022 L 18595 3675 0 3675 Z M 1532 3675 L 1532 1815 20000 1815 20000 16252 C 19298 16252 18595 16352 18595 16352 L 18595 3675 Z M 2972 1815 L 2972 0 21600 0 21600 14392 C 20800 14392 20000 14467 20000 14467 L 20000 1815 Z N F M 0 3675 L 18595 3675 18595 18022 C 9298 18022 9298 23542 0 20782 Z M 1532 3675 L 1532 1815 20000 1815 20000 16252 C 19298 16252 18595 16352 18595 16352 M 2972 1815 L 2972 0 21600 0 21600 14392 C 20800 14392 20000 14467 20000 14467 N F S M 0 20782 C 9298 23542 9298 18022 18595 18022 L 18595 16352 C 18595 16352 19298 16252 20000 16252 L 20000 14467 C 20000 14467 20800 14392 21600 14392 L 21600 0 2972 0 2972 1815 1532 1815 1532 3675 0 3675 Z N')">
        <xsl:value-of select ="'Flowchart: Multi document '"/>
      </xsl:when>
      <!-- Flow Chart: Terminator -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-terminator'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 3475 0 L 18125 0 A ?f59 ?f60 ?f61 ?f62 18125 0 ?f56 ?f58  W ?f63 ?f64 ?f65 ?f66 18125 0 ?f56 ?f58 L 3475 21600 A ?f102 ?f103 ?f104 ?f105 3475 21600 ?f99 ?f101  W ?f106 ?f107 ?f108 ?f109 3475 21600 ?f99 ?f101 Z N')">
        <xsl:value-of select ="'Flowchart: Terminator '"/>
      </xsl:when>
      <!-- Flow Chart: Preparation -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-preparation'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 5 L 2 0 8 0 10 5 8 10 2 10 Z N')">
        <xsl:value-of select ="'Flowchart: Preparation '"/>
      </xsl:when>
      <!-- Flow Chart: Manual-input -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-manual-input'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Manual Input'))">
        <xsl:value-of select ="'Flowchart: Manual Input '"/>
      </xsl:when>
      <!-- Flow Chart: Manual-operation -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-manual-operation'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 0 L 5 0 4 5 1 5 Z N')">
        <xsl:value-of select ="'Flowchart: Manual Operation '"/>
      </xsl:when>
      <!-- Flow Chart: Connector -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-connector'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Connector'))">
        <xsl:value-of select ="'Flowchart: Connector '"/>
      </xsl:when>
      <!-- Flow Chart: Off-page-connector -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-off-page-connector'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Off-page Connector'))">
        <xsl:value-of select ="'Flowchart: Off-page Connector '"/>
      </xsl:when>
      <!-- Flow Chart: Card -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-card'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Card'))">
        <xsl:value-of select ="'Flowchart: Card '"/>
      </xsl:when>
      <!-- Flow Chart: Punched-tape -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-punched-tape'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 2 A ?f63 ?f64 ?f65 ?f66 0 2 ?f60 ?f62  W ?f67 ?f68 ?f69 ?f70 0 2 ?f60 ?f62 A ?f97 ?f98 ?f99 ?f100 ?f60 ?f62 ?f94 ?f96  W ?f101 ?f102 ?f103 ?f104 ?f60 ?f62 ?f94 ?f96 L 20 18 A ?f140 ?f141 ?f142 ?f143 20 18 ?f137 ?f139  W ?f144 ?f145 ?f146 ?f147 20 18 ?f137 ?f139 A ?f170 ?f171 ?f172 ?f173 ?f137 ?f139 ?f167 ?f169  W ?f174 ?f175 ?f176 ?f177 ?f137 ?f139 ?f167 ?f169 Z N')">
        <xsl:value-of select ="'Flowchart: Punched Tape '"/>
      </xsl:when>
      <!-- Flow Chart: Summing-junction -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-summing-junction'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Summing Junction'))">
        <xsl:value-of select ="'Flowchart: Summing Junction '"/>
      </xsl:when>
      <!-- Flow Chart: Or -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-or'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Or'))">
        <xsl:value-of select ="'Flowchart: Or '"/>
      </xsl:when>
      <!-- Flow Chart: Collate -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-collate'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 0 L 2 0 1 1 2 2 0 2 1 1 Z N')">
        <xsl:value-of select ="'Flowchart: Collate '"/>
      </xsl:when>
      <!-- Flow Chart: Sort -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-sort'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Sort'))">
        <xsl:value-of select ="'Flowchart: Sort '"/>
      </xsl:when>
      <!-- Flow Chart: Extract -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-extract'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and starts-with(@draw:name,'Flowchart: Extract'))">
        <xsl:value-of select ="'Flowchart: Extract '"/>
      </xsl:when>
      <!-- Flow Chart: Merge -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-merge'
                     or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 0 L 2 0 1 2 Z N')">
        <xsl:value-of select ="'Flowchart: Merge '"/>
      </xsl:when>
      <!-- Flow Chart: Stored-data -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-stored-data'
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 165100 0 L 990600 0 A ?f57 ?f58 ?f59 ?f60 990600 0 ?f54 ?f56  W ?f61 ?f62 ?f63 ?f64 990600 0 ?f54 ?f56 L 165100 914400 A ?f104 ?f105 ?f106 ?f107 165100 914400 ?f101 ?f103  W ?f108 ?f109 ?f110 ?f111 165100 914400 ?f101 ?f103 Z N')
                       or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path ='M 292100 0 L 1752600 0 A ?f57 ?f58 ?f59 ?f60 1752600 0 ?f54 ?f56  W ?f61 ?f62 ?f63 ?f64 1752600 0 ?f54 ?f56 L 292100 1143000 A ?f104 ?f105 ?f106 ?f107 292100 1143000 ?f101 ?f103  W ?f108 ?f109 ?f110 ?f111 292100 1143000 ?f101 ?f103 Z N') ">
        <xsl:value-of select ="'Flowchart: Stored Data '"/>
      </xsl:when>
      <!-- Flow Chart: Delay-->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-delay'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f3 ?f5 L ?f12 ?f5 A ?f65 ?f66 ?f67 ?f68 ?f12 ?f5 ?f62 ?f64  W ?f69 ?f70 ?f71 ?f72 ?f12 ?f5 ?f62 ?f64 L ?f3 ?f6 Z N')">
        <xsl:value-of select ="'Flowchart: Delay '"/>
      </xsl:when>
      <!-- Flow Chart: Sequential-access -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-sequential-access'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M ?f12 ?f6 A ?f72 ?f73 ?f74 ?f75 ?f12 ?f6 ?f69 ?f71  W ?f76 ?f77 ?f78 ?f79 ?f12 ?f6 ?f69 ?f71 A ?f115 ?f116 ?f117 ?f118 ?f69 ?f71 ?f112 ?f114  W ?f119 ?f120 ?f121 ?f122 ?f69 ?f71 ?f112 ?f114 A ?f158 ?f159 ?f160 ?f161 ?f112 ?f114 ?f155 ?f157  W ?f162 ?f163 ?f164 ?f165 ?f112 ?f114 ?f155 ?f157 A ?f205 ?f206 ?f207 ?f208 ?f155 ?f157 ?f202 ?f204  W ?f209 ?f210 ?f211 ?f212 ?f155 ?f157 ?f202 ?f204 L ?f4 ?f26 ?f4 ?f6 Z N')">
        <xsl:value-of select ="'Flowchart: Sequential Access Storage '"/>
      </xsl:when>
      <!-- Flow Chart: Direct-access-storage -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-direct-access-storage'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M 190500 0 L 952500 0 A ?f58 ?f59 ?f60 ?f61 952500 0 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 952500 0 ?f55 ?f57 L 190500 685800 A ?f101 ?f102 ?f103 ?f104 190500 685800 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 190500 685800 ?f98 ?f100 Z N F M 952500 685800 A ?f113 ?f102 ?f114 ?f104 952500 685800 ?f112 ?f100  W ?f115 ?f106 ?f116 ?f108 952500 685800 ?f112 ?f100 N F M 190500 0 L 952500 0 A ?f58 ?f59 ?f60 ?f61 952500 0 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 952500 0 ?f55 ?f57 L 190500 685800 A ?f101 ?f102 ?f103 ?f104 190500 685800 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 190500 685800 ?f98 ?f100 Z N')
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M 381000 0 L 1905000 0 A ?f58 ?f59 ?f60 ?f61 1905000 0 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 1905000 0 ?f55 ?f57 L 381000 1295400 A ?f101 ?f102 ?f103 ?f104 381000 1295400 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 381000 1295400 ?f98 ?f100 Z N F M 1905000 1295400 A ?f113 ?f102 ?f114 ?f104 1905000 1295400 ?f112 ?f100  W ?f115 ?f106 ?f116 ?f108 1905000 1295400 ?f112 ?f100 N F M 381000 0 L 1905000 0 A ?f58 ?f59 ?f60 ?f61 1905000 0 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 1905000 0 ?f55 ?f57 L 381000 1295400 A ?f101 ?f102 ?f103 ?f104 381000 1295400 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 381000 1295400 ?f98 ?f100 Z N')
                      or  (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path ='S M 250033 0 L 1250165 0 A ?f58 ?f59 ?f60 ?f61 1250165 0 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 1250165 0 ?f55 ?f57 L 250033 857256 A ?f101 ?f102 ?f103 ?f104 250033 857256 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 250033 857256 ?f98 ?f100 Z N F M 1250165 857256 A ?f113 ?f102 ?f114 ?f104 1250165 857256 ?f112 ?f100  W ?f115 ?f106 ?f116 ?f108 1250165 857256 ?f112 ?f100 N F M 250033 0 L 1250165 0 A ?f58 ?f59 ?f60 ?f61 1250165 0 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 1250165 0 ?f55 ?f57 L 250033 857256 A ?f101 ?f102 ?f103 ?f104 250033 857256 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 250033 857256 ?f98 ?f100 Z N')">
        <xsl:value-of select ="'Flowchart: Direct Access Storage'"/>
      </xsl:when>
      <!-- Flow Chart: Magnetic-disk --> 
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-magnetic-disk'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M 0 139700 A ?f58 ?f59 ?f60 ?f61 0 139700 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 139700 ?f55 ?f57 L 762000 698500 A ?f101 ?f102 ?f103 ?f104 762000 698500 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 762000 698500 ?f98 ?f100 Z N F M 762000 139700 A ?f101 ?f113 ?f103 ?f114 762000 139700 ?f98 ?f112  W ?f105 ?f115 ?f107 ?f116 762000 139700 ?f98 ?f112 N F M 0 139700 A ?f58 ?f59 ?f60 ?f61 0 139700 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 139700 ?f55 ?f57 L 762000 698500 A ?f101 ?f102 ?f103 ?f104 762000 698500 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 762000 698500 ?f98 ?f100 Z N')
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path ='S M 0 241300 A ?f58 ?f59 ?f60 ?f61 0 241300 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 241300 ?f55 ?f57 L 1143000 1206500 A ?f101 ?f102 ?f103 ?f104 1143000 1206500 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 1143000 1206500 ?f98 ?f100 Z N F M 1143000 241300 A ?f101 ?f113 ?f103 ?f114 1143000 241300 ?f98 ?f112  W ?f105 ?f115 ?f107 ?f116 1143000 241300 ?f98 ?f112 N F M 0 241300 A ?f58 ?f59 ?f60 ?f61 0 241300 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 241300 ?f55 ?f57 L 1143000 1206500 A ?f101 ?f102 ?f103 ?f104 1143000 1206500 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 1143000 1206500 ?f98 ?f100 Z N')
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path ='S M 0 215900 A ?f58 ?f59 ?f60 ?f61 0 215900 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 215900 ?f55 ?f57 L 990600 1079500 A ?f101 ?f102 ?f103 ?f104 990600 1079500 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 990600 1079500 ?f98 ?f100 Z N F M 990600 215900 A ?f101 ?f113 ?f103 ?f114 990600 215900 ?f98 ?f112  W ?f105 ?f115 ?f107 ?f116 990600 215900 ?f98 ?f112 N F M 0 215900 A ?f58 ?f59 ?f60 ?f61 0 215900 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 215900 ?f55 ?f57 L 990600 1079500 A ?f101 ?f102 ?f103 ?f104 990600 1079500 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 990600 1079500 ?f98 ?f100 Z N')
                      or  (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path ='S M 0 178595 A ?f58 ?f59 ?f60 ?f61 0 178595 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 178595 ?f55 ?f57 L 1214446 892975 A ?f101 ?f102 ?f103 ?f104 1214446 892975 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 1214446 892975 ?f98 ?f100 Z N F M 1214446 178595 A ?f101 ?f113 ?f103 ?f114 1214446 178595 ?f98 ?f112  W ?f105 ?f115 ?f107 ?f116 1214446 178595 ?f98 ?f112 N F M 0 178595 A ?f58 ?f59 ?f60 ?f61 0 178595 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 0 178595 ?f55 ?f57 L 1214446 892975 A ?f101 ?f102 ?f103 ?f104 1214446 892975 ?f98 ?f100  W ?f105 ?f106 ?f107 ?f108 1214446 892975 ?f98 ?f100 Z N')">
        <xsl:value-of select ="'Flowchart: Magnetic Disk '"/>
      </xsl:when>
      <!-- Flow Chart: Display -->
      <xsl:when test ="draw:enhanced-geometry/@draw:type='flowchart-display'
                      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'M 0 457200 L 177800 0 889000 0 A ?f58 ?f59 ?f60 ?f61 889000 0 ?f55 ?f57  W ?f62 ?f63 ?f64 ?f65 889000 0 ?f55 ?f57 L 177800 914400 Z N')">
        <xsl:value-of select ="'Flowchart: Display '"/>
      </xsl:when>
		
      
      
      
		<!-- Action Buttons Back or Previous -->
		<xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt194' and 
						           draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f8 L ?f14 ?f12 ?f14 ?f16 Z N')
                        or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z M ?f14 ?f6 L ?f15 ?f12 ?f15 ?f13 Z N S M ?f14 ?f6 L ?f15 ?f12 ?f15 ?f13 Z N F M ?f14 ?f6 L ?f15 ?f12 ?f15 ?f13 Z N F M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z N')">
			<xsl:value-of select ="'actionButtonBackPrevious'"/>
		</xsl:when>
		<!-- Action Buttons Forward or Next -->
		<xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt193' and 
						draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f12 L ?f14 ?f8 ?f10 ?f16 Z N')
            or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z M ?f15 ?f6 L ?f14 ?f12 ?f14 ?f13 Z N S M ?f15 ?f6 L ?f14 ?f12 ?f14 ?f13 Z N F M ?f15 ?f6 L ?f14 ?f13 ?f14 ?f12 Z N F M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z N')">
			<xsl:value-of select ="'actionButtonForwardNext'"/>
		</xsl:when>
		<!-- Action Buttons Beginning -->
		<xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt196') and 
						(draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f8 L ?f14 ?f12 ?f14 ?f16 Z N M ?f18 ?f12 L ?f20 ?f12 ?f20 ?f16 ?f18 ?f16 Z N')">
			<xsl:value-of select ="'actionButtonBeginning'"/>
		</xsl:when>
		<!-- Action Buttons end -->
		<xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt195') and 
						(draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f22 ?f8 L ?f18 ?f16 ?f18 ?f12 Z N M ?f24 ?f12 L ?f24 ?f16 ?f14 ?f16 ?f14 ?f12 Z N')">
			<xsl:value-of select ="'actionButtonEnd'"/>
		</xsl:when>
		<!-- Action Buttons Home -->
		<xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt190') and 
						(draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f7 ?f10 L ?f12 ?f14 ?f12 ?f16 ?f18 ?f16 ?f18 ?f20 ?f22 ?f8 ?f24 ?f8 ?f24 ?f26 ?f28 ?f26 ?f28 ?f8 ?f30 ?f8 Z N M ?f12 ?f14 L ?f12 ?f16 ?f18 ?f16 ?f18 ?f20 Z N M ?f32 ?f36 L ?f34 ?f36 ?f34 ?f26 ?f24 ?f26 ?f24 ?f8 ?f28 ?f8 ?f28 ?f26 ?f32 ?f26 Z N')">
			<xsl:value-of select ="'actionButtonHome'"/>
		</xsl:when>
		<!-- Action Buttons Information -->
		<xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt192' and 
						draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f7 ?f12 X ?f10 ?f8 ?f7 ?f16 ?f14 ?f8 ?f7 ?f12 Z N M ?f7 ?f20 X ?f18 ?f42 ?f7 ?f24 ?f22 ?f42 ?f7 ?f20 Z N M ?f26 ?f28 L ?f30 ?f28 ?f30 ?f32 ?f34 ?f32 ?f34 ?f36 ?f26 ?f36 ?f26 ?f32 ?f38 ?f32 ?f38 ?f40 ?f26 ?f40 Z N')
      or (draw:enhanced-geometry/@draw:type='non-primitive' and draw:enhanced-geometry/@draw:enhanced-path = 'S M ?f3 ?f5 L ?f4 ?f5 ?f4 ?f6 ?f3 ?f6 Z M ?f12 ?f16 A ?f76 ?f77 ?f78 ?f79 ?f12 ?f16 ?f73 ?f75  W ?f80 ?f81 ?f82 ?f83 ?f12 ?f16 ?f73 ?f75 Z N S M ?f12 ?f16 A ?f76 ?f77 ?f78 ?f79 ?f12 ?f16 ?f73 ?f75  W ?f80 ?f81 ?f82 ?f83 ?f12 ?f16 ?f73 ?f75 Z M ?f12 ?f27 A ?f104 ?f105 ?f106 ?f107 ?f12 ?f27 ?f101 ?f103  W ?f108 ?f109 ?f110 ?f111 ?f12 ?f27 ?f101 ?f103 M ?f32 ?f28 L ?f32 ?f29 ?f33 ?f29 ?f33 ?f30 ?f32 ?f30 ?f32 ?f31 ?f35 ?f31 ?f35 ?f30 ?f34 ?f30 ?f34 ?f28 Z N S M ?f12 ?f27 A ?f104 ?f105 ?f106 ?f107 ?f12 ?f27 ?f101 ?f103  W ?f108 ?f109 ?f110 ?f111 ?f12 ?f27 ?f101 ?f103 M ?f32 ?f28 L ?f34 ?f28 ?f34 ?f30 ?f35 ?f30 ?f35 ?f31 ?f32 ?f31 ?f32 ?f30 ?f33 ?f30 ?f33 ?f29 ?f32 ?f29 Z N F M ?f12 ?f16 A ?f76 ?f77 ?f78 ?f79 ?f12 ?f16 ?f73 ?f75  W ?f80 ?f81 ?f82 ?f83 ?f12 ?f16 ?f73 ?f75 Z M ?f12 ?f27 A ?f104 ?f105 ?f106 ?f107 ?f12 ?f27 ?f101 ?f103  W ?f108 ?f109 ?f110 ?f111 ?f12 ?f27 ?f101 ?f103 M ?f32 ?f28 L ?f34 ?f28 ?f34 ?f30 ?f35 ?f30 ?f35 ?f31 ?f32 ?f31 ?f32 ?f30 ?f33 ?f30 ?f33 ?f29 ?f32 ?f29 Z N F M ?f3 ?f5 L ?f4 ?f5 ?f4 ?f6 ?f3 ?f6 Z N')">
			<xsl:value-of select ="'actionButtonInformation'"/>
		</xsl:when>
		<!-- Action Buttons Return -->
		<xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt197') and 
						(draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f12 L ?f14 ?f12 ?f14 ?f16 C ?f14 ?f18 ?f20 ?f22 ?f24 ?f22 L ?f7 ?f22 C ?f26 ?f22 ?f28 ?f18 ?f28 ?f16 L ?f28 ?f12 ?f7 ?f12 ?f30 ?f32 ?f34 ?f12 ?f36 ?f12 ?f36 ?f16 C ?f36 ?f38 ?f40 ?f42 ?f7 ?f42 L ?f24 ?f42 C ?f44 ?f42 ?f10 ?f38 ?f10 ?f16 Z N')">
			<xsl:value-of select ="'actionButtonReturn'"/>
		</xsl:when>		
    </xsl:choose>
  </xsl:template>
 <xsl:template name="tmpSuperSubScriptForward">
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
      <xsl:when test="contains(substring-before(style:text-properties/@style:text-position,' '),'%') = '%'">
        <xsl:variable name="val">
          <xsl:value-of select="number(substring-before(substring-before(style:text-properties/@style:text-position,' '),'%'))"/>
        </xsl:variable>
        <xsl:attribute name="baseline">
          <xsl:value-of select="$val * 1000"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpSMShapeFillColor">
    <xsl:param name ="shapeCount" />
    <xsl:param name="grpFlag" />
    <xsl:param name ="UniqueId" />
    <xsl:choose>
      <xsl:when test="@draw:fill='solid'">
        <xsl:if test="@draw:fill-color">
          <a:solidFill>
						<xsl:variable name="varSrgbVal">
							<xsl:value-of select="substring-after(@draw:fill-color,'#')"/>
						</xsl:variable>
						<xsl:if test="$varSrgbVal != ''">
            <a:srgbClr>
              <xsl:attribute name="val">
									<xsl:value-of select="$varSrgbVal" />
              </xsl:attribute>
              <xsl:if test ="@draw:opacity">
                <xsl:variable name="tranparency" select="substring-before(@draw:opacity,'%')"/>
                <xsl:call-template name="tmpshapeTransperancy">
                  <xsl:with-param name="tranparency" select="$tranparency"/>
                </xsl:call-template>
              </xsl:if>
            </a:srgbClr>
						</xsl:if>
          </a:solidFill>
        </xsl:if>
      </xsl:when>
      <xsl:when test="@draw:fill='none'">
        <a:noFill/>
      </xsl:when>
      <xsl:when test="@draw:fill='gradient'">
        <xsl:call-template name="tmpGradientFill">
          <xsl:with-param name="gradStyleName" select="@draw:fill-gradient-name"/>
          <xsl:with-param  name="opacity" select="substring-before(@draw:opacity,'%')"/>
        </xsl:call-template>
      </xsl:when>

      <!--Added by Mathi-->
      <xsl:when test="(@draw:fill='bitmap') and $grpFlag!='true'">
        <xsl:call-template name="tmpBitmapFill">
          <xsl:with-param name="FileName" select="concat('bitmap',$shapeCount)" />
          <xsl:with-param name="var_imageName" select="@draw:fill-image-name"/>
          <xsl:with-param  name="opacity" select="substring-before(@draw:opacity,'%')"/>
          <xsl:with-param  name="stretch" select="style:graphic-properties/@style:repeat"/>
              </xsl:call-template>
            </xsl:when>
      <xsl:when test="(@draw:fill='bitmap') and $grpFlag='true'">
        <xsl:call-template name="tmpBitmapFill">
          <xsl:with-param name="FileName" select="concat('grpbitmap',$UniqueId)" />
          <xsl:with-param name="var_imageName" select="@draw:fill-image-name" />
          <xsl:with-param name ="UniqueId" select ="$UniqueId" />
          <xsl:with-param  name="opacity" select="substring-before(@draw:opacity,'%')"/>
          <xsl:with-param  name="stretch" select="style:graphic-properties/@style:repeat"/>
        </xsl:call-template>
      </xsl:when>

    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpInternalPadding">
    <xsl:variable name="UnitTop">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="@fo:padding-top"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="UnitBottom">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="@fo:padding-bottom"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="UnitLeft">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="@fo:padding-left"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="UnitRight">
      <xsl:call-template name="getConvertUnit">
        <xsl:with-param name="length" select="@fo:padding-right"/>
      </xsl:call-template>
    </xsl:variable>

       <xsl:choose>
      <xsl:when test ="@fo:padding-top and
					substring-before(@fo:padding-top,$UnitTop) &gt; 0">
        <xsl:attribute name ="tIns">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@fo:padding-top"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        </xsl:when>
      <xsl:when test ="substring-before(@fo:padding,$UnitTop) = '0' or substring-before(@fo:padding,$UnitTop) = ''">
        <xsl:attribute name ="tIns">
          <xsl:value-of select ="0"/>
        </xsl:attribute>
      </xsl:when >
    </xsl:choose>
       <xsl:choose>
      <xsl:when test="@fo:padding-bottom and
					substring-before(@fo:padding-bottom,$UnitBottom) &gt; 0">
        <xsl:attribute name ="bIns">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@fo:padding-bottom"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        </xsl:when>
      <xsl:when test ="substring-before(@fo:padding,$UnitBottom) = '0' or substring-before(@fo:padding,$UnitBottom) = ''">
          <xsl:attribute name ="bIns">
            <xsl:value-of select ="0"/>
          </xsl:attribute>
        </xsl:when >
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="@fo:padding-left and
					substring-before(@fo:padding-left,$UnitLeft) &gt; 0">
          <xsl:attribute name ="lIns">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@fo:padding-left"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test ="substring-before(@fo:padding,$UnitLeft) = '0' or substring-before(@fo:padding,$UnitLeft) = ''">
          <xsl:attribute name ="lIns">
          <xsl:value-of select ="0"/>
          </xsl:attribute>
        </xsl:when >
      </xsl:choose>
      <xsl:choose>
      <xsl:when test="@fo:padding-right and
					substring-before(@fo:padding-right,$UnitRight) &gt; 0">
        <xsl:attribute name ="rIns">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name="length"  select ="@fo:padding-right"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
        <xsl:when test ="substring-before(@fo:padding,$UnitRight) = '0' or substring-before(@fo:padding,$UnitRight) = ''">
        <xsl:attribute name ="rIns">
          <xsl:value-of select ="0"/>
        </xsl:attribute>
        </xsl:when >
         </xsl:choose>
  </xsl:template>

	<!--CalloutAdjs Template (added by Mathi)-->
	<xsl:template name="tmpCalloutAdjustment">
		<xsl:param name="prst"/>
		<xsl:variable name="convertUnit">
			<xsl:call-template name="getConvertUnit">
				<xsl:with-param name="length" select="@draw:transform"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="width" >
			<xsl:call-template name="convertUnitsToCm">
				<xsl:with-param name="length" select="@svg:width"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="height" >
			<xsl:call-template name="convertUnitsToCm">
				<xsl:with-param name="length" select="@svg:height"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="angle">
			<xsl:choose>
				<xsl:when test="@draw:transform">
          <xsl:call-template name="tmpGetGroupRotXYVals">
            <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
            <xsl:with-param name="XY" select="'ROT'"/>
          </xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name ="x">
			<xsl:choose>
				<xsl:when test="@draw:transform">

          <xsl:variable name="tmp_X">
            <xsl:call-template name="tmpGetGroupRotXYVals">
              <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
              <xsl:with-param name="unit" select="$convertUnit"/>
              <xsl:with-param name="XY" select="'X'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length" select="concat($tmp_X,$convertUnit)" />
          </xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name="convertUnitsToCm">
						<xsl:with-param name="length" select="@svg:x"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name ="y">
			<xsl:choose>
				<xsl:when test="@draw:transform">
          <xsl:variable name="tmp_Y">
            <xsl:call-template name="tmpGetGroupRotXYVals">
              <xsl:with-param name="drawTranformVal" select="@draw:transform"/>
              <xsl:with-param name="unit" select="$convertUnit"/>
              <xsl:with-param name="XY" select="'Y'"/>
            </xsl:call-template>
          </xsl:variable>
          
          <xsl:call-template name="convertUnitsToCm">
            <xsl:with-param name="length" select="concat($tmp_Y,$convertUnit)" />
          </xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name="convertUnitsToCm">
						<xsl:with-param name="length" select="@svg:y"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name ="flipH">
			<xsl:choose>
				<xsl:when test="draw:enhanced-geometry/@draw:mirror-horizontal">
					<xsl:value-of select="(draw:enhanced-geometry/@draw:mirror-horizontal)" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name ="flipV">
			<xsl:choose>
				<xsl:when test="draw:enhanced-geometry/@draw:mirror-vertical">
					<xsl:value-of select="(draw:enhanced-geometry/@draw:mirror-vertical)" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
    <xsl:if test ="($prst='rectangular-callout'
		                or $prst='wedgeRectCallout'
						or $prst='wedgeRoundRectCallout'
                                                or $prst='wedgeEllipseCallout'
						or $prst='cloudCallout')">
			<a:prstGeom>
				<xsl:attribute name="prst">
					<xsl:value-of select="$prst"/>
				</xsl:attribute>
				<a:avLst>
					<a:gd name="adj1">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectAdj1Notlinefmla1:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj2">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectAdj2Notlinefmla2:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<xsl:if test="$prst='wedgeRoundRectCallout'">
						<a:gd name="adj3" fmla="val 16667" />
					</xsl:if>
				</a:avLst>
			</a:prstGeom>
		</xsl:if>
		<xsl:if test ="(draw:enhanced-geometry/@draw:type='line-callout-1' or draw:enhanced-geometry/@draw:type='line-callout-3')">
			<a:prstGeom prst="borderCallout1">
				<a:avLst>
					<a:gd name="adj1">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine1Adj1fmla1:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj2">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine1Adj2fmla2:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj3">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine1Adj3fmla3:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj4">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine1Adj4fmla4:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
				</a:avLst>
			</a:prstGeom>
		</xsl:if>
		<xsl:if test ="(draw:enhanced-geometry/@draw:type='line-callout-2')">
			<a:prstGeom prst="borderCallout2">
				<a:avLst>
					<a:gd name="adj1">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine2Adj1fmla1:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj2">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine2Adj2fmla2:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj3">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine2Adj3fmla3:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj4">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine2Adj4fmla4:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj5">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine2Adj5fmla5:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj6">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine2Adj6fmla6:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
				</a:avLst>
			</a:prstGeom>
		</xsl:if>
		<xsl:if test ="(draw:enhanced-geometry/@draw:type='mso-spt49') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 Z N M ?f6 ?f7 F L ?f4 ?f5 ?f2 ?f3 ?f0 ?f1 N')">
			<a:prstGeom prst="borderCallout3">
				<a:avLst>
					<a:gd name="adj1">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj1fmla1:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj2">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj2fmla2:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj3">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj3fmla3:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj4">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj4fmla4:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj5">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj5fmla5:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj6">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj6fmla6:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj7">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj7fmla7:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
					<a:gd name="adj8">
						<xsl:attribute name="fmla">
							<xsl:value-of select="concat('Callout-DirectLine3Adj8fmla8:',$width,':',$height,':',$x,':',$y,':',$flipH,':',$flipV,':',$angle,'@',draw:enhanced-geometry/@draw:modifiers)"/>
						</xsl:attribute>
					</a:gd>
				</a:avLst>
			</a:prstGeom>
		</xsl:if>
	</xsl:template>
	
  <xsl:template name="tmpHyperLnkBuImgRel">
    <xsl:param name="var_pos"/>
    <xsl:param name="shapeId"/>
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
    <xsl:for-each select ="node()">
      <xsl:if test ="name()='text:p'" >
        <xsl:if test="text:a/@xlink:href !=''">
          <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <xsl:attribute name="Id">
              <xsl:value-of select="concat($shapeId,'Link',position())"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
                <xsl:attribute name="Type">
                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:') or contains(.,'https://')]">
                <xsl:attribute name="Type">
                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:value-of select="text:a/@xlink:href"/>
                </xsl:attribute>
                <xsl:attribute name="TargetMode">
                  <xsl:value-of select="'External'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
                  </xsl:attribute>
                  <xsl:attribute name="TargetMode">
                    <xsl:value-of select="'External'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <!--links Absolute Path-->
                    <xsl:variable name ="xlinkPath" >
                      <xsl:value-of select ="text:a/@xlink:href"/>
                    </xsl:variable>
                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
                  </xsl:attribute>
                  <xsl:attribute name="TargetMode">
                    <xsl:value-of select="'External'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
          </Relationship>
        </xsl:if>
        <xsl:if test="text:span/text:a/@xlink:href !=''">
          <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <xsl:attribute name="Id">
              <xsl:value-of select="concat($shapeId,'Link',position())"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
                <xsl:attribute name="Type">
                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="Type">
                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:value-of select="text:span/text:a/@xlink:href"/>
                </xsl:attribute>
                <xsl:attribute name="TargetMode">
                  <xsl:value-of select="'External'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
                  </xsl:attribute>
                  <xsl:attribute name="TargetMode">
                    <xsl:value-of select="'External'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <!--links Absolute Path-->
                    <xsl:variable name ="xlinkPath" >
                      <xsl:value-of select ="text:span/text:a/@xlink:href"/>
                    </xsl:variable>
                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
                  </xsl:attribute>
                  <xsl:attribute name="TargetMode">
                    <xsl:value-of select="'External'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
          </Relationship>
        </xsl:if>
      </xsl:if>
      <xsl:if test ="name()='text:list'" >
        <xsl:variable name ="listId">
          <xsl:choose>
            <xsl:when test="contains($shapeId,'text-box')">
              <xsl:value-of select ="./@text:style-name"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="./@text:style-name"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:call-template name="tmpBulletImageRel">
          <xsl:with-param name ="var_pos" select="$var_pos" />
          <xsl:with-param name ="shapeId" select="$shapeId" />
          <xsl:with-param name ="listId" select="$listId" />
          <xsl:with-param name ="grpFlag" select="$grpFlag" />
          <xsl:with-param name ="UniqueId" select="$UniqueId" />
          <xsl:with-param name ="listItemCount" select="generate-id()" />          
        </xsl:call-template>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpOLEObjects">
    <xsl:param name ="pageNo" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
    <xsl:choose>
      <xsl:when test="document(concat(substring-after(./child::node()[1]/@xlink:href,'./'),'/content.xml'))/child::node() or
                      document(concat(translate(./child::node()[1]/@xlink:href,'/',''),'/content.xml'))/child::node() ">
        <pzip:copy pzip:source="#CER#WordprocessingConverter.dll#OdfConverter.Wordprocessing.resources.OLEplaceholder.png#"
              pzip:target="{concat('ppt/media/','oleObjectImage_',generate-id(),'.png')}"/>
        <p:pic>
          <p:nvPicPr>
            <p:cNvPr id="{$shapeCount + 1 }">
              <xsl:attribute name="name">
                <xsl:value-of select="concat('Picture ',$shapeCount+1)"/>
              </xsl:attribute>
              <xsl:attribute name="descr">
                <xsl:value-of select="concat('oleObjectImage_',generate-id(),'.png')"/>
              </xsl:attribute>
            </p:cNvPr >
            <p:cNvPicPr>
              <a:picLocks noChangeAspect="1">
                <xsl:choose>
                  <xsl:when test="$grpFlag!='true'">
                    <xsl:attribute name="noGrp">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </a:picLocks>
            </p:cNvPicPr>
            <p:nvPr/>
          </p:nvPicPr>
          <p:blipFill>
            <a:blip>
              <xsl:attribute name ="r:embed">
                <xsl:value-of select ="concat('oleObjectImage_',generate-id())"/>
              </xsl:attribute>
            </a:blip >
            <a:stretch>
              <a:fillRect />
            </a:stretch>
          </p:blipFill>
          <p:spPr>
            <xsl:for-each select=".">
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
            </xsl:for-each>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
          </p:spPr>
        </p:pic>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="tmpOLE">
          <xsl:with-param name ="pageNo" select="$pageNo" />
          <xsl:with-param name ="shapeCount" select="$shapeCount" />
          <xsl:with-param name ="grpFlag" select="$grpFlag" />
          <xsl:with-param name ="UniqueId" select="$UniqueId" />
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOLE">
    <xsl:param name ="pageNo" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
    <xsl:param name ="OLEasXML" />
    <xsl:variable name="OLEasXMLImage">
      <xsl:value-of select="document(concat(substring-after(./child::node()[1]/@xlink:href,'./'),'/content.xml'))
                           /office:document-content/office:body/office:drawing/draw:page/draw:frame/draw:image/@xlink:href"/>
    </xsl:variable>
    
        <xsl:for-each select="./child::node()[1]">
          <xsl:if test="@xlink:href !=''">
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr>
                <xsl:attribute name="id">
                  <xsl:value-of select="$shapeCount+1"/>
                </xsl:attribute>
                <xsl:attribute name="name">
                  <xsl:value-of select="concat('Object ',$shapeCount+1)"/>
                </xsl:attribute>
              </p:cNvPr>
              <p:cNvGraphicFramePr>
                <a:graphicFrameLocks noChangeAspect="1"/>
              </p:cNvGraphicFramePr>
              <p:nvPr/>
            </p:nvGraphicFramePr>
            <xsl:for-each select="..">
              <xsl:call-template name ="tmpdrawCordinates">
                <xsl:with-param name="OLETAB" select="'true'"/>
                <xsl:with-param name="grpFlag" select="$grpFlag"/>
              </xsl:call-template>
            </xsl:for-each>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
                <p:oleObj showAsIcon="1" imgW="2790476" imgH="533474">
                  <xsl:attribute name="spid">
                                    <xsl:value-of select="concat('_x0000_s', $pageNo * 1024 + $shapeCount)"/>
                                  </xsl:attribute>
                    <xsl:attribute name="name">
                    <xsl:value-of select="'Package'"/>
                  </xsl:attribute>
                  <xsl:attribute name="r:id">
                    <xsl:choose>
                      <xsl:when test="$grpFlag='true'">
                        <xsl:value-of select="concat('Slidegrp',$pageNo,'_Ole',generate-id())"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="concat('Slide',$pageNo,'_Ole',generate-id())"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                  <xsl:attribute name="progId">
                    <!--<xsl:value-of select="'Package'"/>-->
                    <xsl:choose>
                      <xsl:when test="contains(@xlink:href,'.pptx')">
                        <xsl:value-of select="'PowerPoint.Show.12'"/>
                      </xsl:when>
                      <xsl:when test="contains(@xlink:href,'.xslx')">
                        <xsl:value-of select="'Excel.Sheet.12'"/>
                      </xsl:when>
                      <xsl:when test="contains(@xlink:href,'.xsl')">
                        <xsl:value-of select="'Excel.Sheet.8'"/>
                      </xsl:when>
                      <xsl:when test="contains(@xlink:href,'.ppt')">
                        <xsl:value-of select="'PowerPoint.Show.8'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'Package'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                  <xsl:choose>
                    <xsl:when test="starts-with(@xlink:href,'./')">
                      <p:embed/>
                    </xsl:when>
                      <xsl:when test="starts-with(@xlink:href,'/') or starts-with(@xlink:href,'../') or starts-with(@xlink:href,'//')
                                            or starts-with(@xlink:href,'file:///')">
                      <p:link/>
                    </xsl:when>
                  <xsl:otherwise>
                    <p:embed/>
                  </xsl:otherwise>
                  </xsl:choose>
                </p:oleObj>
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
          <xsl:variable name="extension">
            <xsl:variable name="objectName">
              <xsl:choose>
                <xsl:when test="starts-with(@xlink:href,'./')">
                  <xsl:value-of select="substring-after(@xlink:href,'./')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@xlink:href"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="substring-after($objectName,'.')!=''">
                <xsl:value-of select="concat('.',substring-after($objectName,'.'))"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="starts-with(@xlink:href,'./')">
              <xsl:variable name="target">
                <xsl:choose>
                  <xsl:when test="$extension!=''">
                    <xsl:value-of select="concat('ppt/embeddings/',translate(substring-after(@xlink:href,'./'),' ',''))"/>
                  </xsl:when>
                  <xsl:when test="$extension=''">
                    <xsl:value-of select="concat('ppt/embeddings/',translate(substring-after(@xlink:href,'./'),' ',''),'.bin')"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:variable>
            <xsl:choose>
              <xsl:when test="$OLEasXML='true'">
                <pzip:copy pzip:source="{concat(substring-after(@xlink:href,'./'),'/',$OLEasXMLImage)}" pzip:target="{$target}" />
                
              </xsl:when>
              <xsl:otherwise>
              <pzip:copy pzip:source="{substring-after(@xlink:href,'./')}" pzip:target="{$target}" />
              </xsl:otherwise>
            </xsl:choose>
            
            </xsl:when>
            <xsl:when test="starts-with(@xlink:href,'/')"></xsl:when>
            <xsl:when test="starts-with(@xlink:href,'\\') or starts-with(@xlink:href,'//')"></xsl:when>
          <xsl:otherwise>
						<!--<pzip:copy pzip:source="{translate(@xlink:href,'/','')}" pzip:target="{concat('ppt/embeddings/',translate(translate(@xlink:href,' ',''),'/',''),'.bin')}" />-->
						<pzip:copy pzip:source="{translate(@xlink:href,'/','')}" pzip:target="{concat('ppt/embeddings/',translate(translate(translate(@xlink:href,' ',''),'/',''),':','_'),'.bin')}" />
          </xsl:otherwise>
          </xsl:choose>
          <xsl:if test="./parent::node()/draw:image/@xlink:href !=''">
            <xsl:choose>
              <xsl:when test="starts-with(./parent::node()/draw:image/@xlink:href,'./')">
                <xsl:variable name="olePictureType">
                  <xsl:call-template name="GetOLEPictureType">
                    <xsl:with-param name="olePicture" select="substring-after(./parent::node()/draw:image/@xlink:href,'./')" />
                    <xsl:with-param name="oleType">
                      <xsl:choose>
                        <xsl:when test="starts-with(@xlink:href,'./')">
                          <xsl:value-of select="'embed'"/>
                        </xsl:when>
                        <xsl:when test="starts-with(@xlink:href,'/') or starts-with(@xlink:href,'../') or starts-with(@xlink:href,'//')
                                            or starts-with(@xlink:href,'file:///')">
                          <xsl:value-of select="'link'"/>
                        </xsl:when>
                      </xsl:choose>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="$olePictureType='GDIMetaFile'">
                    <pzip:copy pzip:source="#CER#WordprocessingConverter.dll#OdfConverter.Wordprocessing.resources.OLEplaceholder.png#"
                 pzip:target="{concat('ppt/media/','oleObjectImage_',generate-id(),'.png')}"/>
                  </xsl:when>
                  <xsl:otherwise>
                <pzip:copy   pzip:source="{substring-after(./parent::node()/draw:image/@xlink:href,'./')}"
                   pzip:target="{concat('ppt/media/','oleObjectImage_',generate-id(),'.png')}" />
                  </xsl:otherwise>
                </xsl:choose>

              </xsl:when>
            <xsl:otherwise>
              <pzip:copy   pzip:source="{./parent::node()/draw:image/@xlink:href}"
                  pzip:target="{concat('ppt/media/','oleObjectImage_',generate-id(),'.png')}" />
            </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
          </xsl:if>
        </xsl:for-each>      
  </xsl:template>
  <xsl:template name="GetOLEPictureType">
    <xsl:param name="olePicture" />
    <xsl:param name="oleType" />
    
    <xsl:variable name="allManifestEntries" select="document('META-INF/manifest.xml')/manifest:manifest/manifest:file-entry" />
    <xsl:variable name="type" select="$allManifestEntries[@manifest:full-path=$olePicture]/@manifest:media-type" />

    <xsl:choose>
      <xsl:when test="contains($type,'application/x-openoffice-gdimetafile')">
        <xsl:text>GDIMetaFile</xsl:text>
      </xsl:when>
      <xsl:when test="$type=''">
        <xsl:if test="$oleType='link'">
        <xsl:text>GDIMetaFile</xsl:text>
        </xsl:if>
      </xsl:when>
      <!-- picture is a WMF -->
      <xsl:when test="$type='application/x-openoffice-wmf;windows_formatname=&quot;Image WMF&quot;'">
        <xsl:text>wmf</xsl:text>
      </xsl:when>
      <!-- picture is a PNG -->
      <xsl:when test="$type='image/png'">
        <xsl:text>png</xsl:text>
      </xsl:when>
      <!-- picture is a JPG -->
      <xsl:when test="$type='image/jpeg'">
        <xsl:text>jpeg</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>png</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

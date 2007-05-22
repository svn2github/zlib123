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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">

  <xsl:import href="cell.xsl"/>
  <xsl:import href="common.xsl"/>

  <xsl:template name="InsertDrawing">
    <xdr:wsDr>
      <!-- for charts, file pointed by draw:frame/draw:object/@xlink:href must contain office:chart -->
      <xsl:for-each
        select="descendant::draw:frame/draw:object[document(concat(translate(@xlink:href,'./',''),'/content.xml'))/office:document-content/office:body/office:chart]">
        <xdr:twoCellAnchor>
          <xdr:from>
            <xdr:col>
              <xsl:call-template name="InsertStartColumn"/>
            </xdr:col>
            <xdr:colOff>
              <xsl:call-template name="InsertStartColumnOffset"/>
            </xdr:colOff>
            <xdr:row>
              <xsl:call-template name="InsertStartRow"/>
            </xdr:row>
            <xdr:rowOff>
            <xsl:call-template name="InsertStartRowOffset"/>
            </xdr:rowOff>
          </xdr:from>
          <xdr:to>
            <xdr:col>
              <xsl:call-template name="InsertEndColumn"/>
            </xdr:col>
            <xdr:colOff>
              <xsl:call-template name="InsertEndColumnOffset"/>
            </xdr:colOff>
            <xdr:row>
              <xsl:call-template name="InsertEndRow"/>
            </xdr:row>
            <xdr:rowOff>
              <xsl:call-template name="InsertEndRowOffset"/>
            </xdr:rowOff>
          </xdr:to>
          <xdr:graphicFrame macro="">
            <xdr:nvGraphicFramePr>
              <xdr:cNvPr id="{position()}" name="{concat('Chart ',position())}"/>
              <xdr:cNvGraphicFramePr>
                <a:graphicFrameLocks/>
              </xdr:cNvGraphicFramePr>
            </xdr:nvGraphicFramePr>
            <xdr:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="0" cy="0"/>
            </xdr:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  r:id="{generate-id(.)}"/>
              </a:graphicData>
            </a:graphic>
          </xdr:graphicFrame>
          <xdr:clientData/>
        </xdr:twoCellAnchor>
      </xsl:for-each>
      
      <xsl:for-each
        select="descendant::draw:frame/draw:image[contains(@xlink:href, 'Pictures')]">

        <xdr:twoCellAnchor editAs="oneCell">
          
          
          <xdr:from>
            <xdr:col>
              <xsl:call-template name="InsertStartColumn"/>
            </xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>
              <xsl:call-template name="InsertStartRow"/>
            </xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
          </xdr:from>
          
          
          
          <xdr:to>
            <xdr:col>13</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>33</xdr:row>
            <xdr:rowOff>114300</xdr:rowOff>
          </xdr:to>
          
          
          <xdr:pic>
            <xdr:nvPicPr>
              <xdr:cNvPr>
                <xsl:attribute name="id">
                  <xsl:value-of select="position()"/>
                </xsl:attribute>
              <xsl:attribute name="name">
                <xsl:value-of select="parent::draw:frame/@draw:name"/>
              </xsl:attribute>
                <xsl:attribute name="descr">
                  <xsl:value-of select="parent::draw:frame/@draw:name"/>
                </xsl:attribute>
              </xdr:cNvPr>
              <xdr:cNvPicPr>
                <a:picLocks noChangeAspect="1"/>
              </xdr:cNvPicPr>
            </xdr:nvPicPr>
            
            <xdr:blipFill>
              <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <xsl:attribute name="r:embed">
                  <xsl:value-of select="generate-id()"/>
                </xsl:attribute>
              </a:blip>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </xdr:blipFill>
            
            <xdr:spPr>
              <a:xfrm>
                <a:off x="5486400" y="4572000"/>
                <a:ext cx="2438400" cy="1828800"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </xdr:spPr>
            
          </xdr:pic>
          <xdr:clientData/>
        </xdr:twoCellAnchor>
        
        
        </xsl:for-each>

    </xdr:wsDr>
    
  </xsl:template>

  <xsl:template name="InsertStartColumn">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <xsl:for-each select="parent::node()/parent::node()">
          <xsl:variable name="position">
            <xsl:value-of select="count(preceding-sibling::table:table-cell) + 1"/>
          </xsl:variable>
          <xsl:variable name="number">
            <xsl:for-each select="parent::node()/table:table-cell[1]">
              <xsl:call-template name="GetColNumber">
                <xsl:with-param name="position" select="$position"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:variable>
          <xsl:value-of select="$number - 1"/>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>1</xsl:text>
      </xsl:when>
    </xsl:choose>    
  </xsl:template>

  <xsl:template name="InsertStartRow">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <!-- get parent table:table-row id -->
        <xsl:variable name="rowId">
          <xsl:value-of select="generate-id(ancestor::table:table-row)"/>
        </xsl:variable>
        <!-- go to first table:table-row-->
        <xsl:for-each select="ancestor::table:table/descendant::table:table-row[1]">
          <xsl:variable name="number">
            <xsl:call-template name="GetRowNumber">
              <xsl:with-param name="rowId" select="$rowId"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="$number - 1"/>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>31</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertEndColumn">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <xsl:for-each select="parent::node()">
          <xsl:variable name="number">
            <xsl:call-template name="GetColNum">
              <xsl:with-param name="cell" select="substring-after(@table:end-cell-address,'.')"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="$number - 1"/>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>5</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="InsertEndRow">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <xsl:for-each select="parent::node()">
          <xsl:variable name="number">
            <xsl:call-template name="GetRowNum">
              <xsl:with-param name="cell" select="substring-after(@table:end-cell-address,'.')"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="$number - 1"/>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>46</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertStartColumnOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <xsl:for-each select="parent::node()">
          <xsl:call-template name="emu-measure">
            <xsl:with-param name="length" select="@svg:x"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>714375</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertStartRowOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <xsl:for-each select="parent::node()">
          <xsl:call-template name="emu-measure">
            <xsl:with-param name="length" select="@svg:y"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>104775</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="InsertEndColumnOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <xsl:for-each select="parent::node()">
          <xsl:call-template name="emu-measure">
            <xsl:with-param name="length" select="@ table:end-x"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>447675</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="InsertEndRowOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:table-cell']">
        <xsl:for-each select="parent::node()">
          <xsl:call-template name="emu-measure">
            <xsl:with-param name="length" select="@ table:end-y"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()/parent::node()[name() = 'table:shapes']">
        <xsl:text>104775</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

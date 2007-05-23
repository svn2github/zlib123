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
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">

  <xsl:import href="cell.xsl"/>
  <xsl:import href="common.xsl"/>

  <!-- Insert Drawing (picture, chart)  -->
  <xsl:template name="InsertDrawing">
    <xdr:wsDr>
      <!--Insert Chart -->
      <xsl:for-each
        select="descendant::draw:frame/draw:object[document(concat(translate(@xlink:href,'./',''),'/content.xml'))/office:document-content/office:body/office:chart]">
        <xdr:twoCellAnchor>
         <xsl:call-template name="SetPosition"/>
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
      
      <!--Insert Picture -->
      <xsl:for-each
        select="descendant::draw:frame/draw:image[contains(@xlink:href, 'Pictures')]">

        <xdr:twoCellAnchor>
          <xsl:call-template name="SetPosition"/>
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
            
           <xsl:call-template name="InsertImageBorders"/>
            
          </xdr:pic>
          <xdr:clientData/>
        </xdr:twoCellAnchor>
        
        
        </xsl:for-each>

    </xdr:wsDr>
    
  </xsl:template>
  
  <!-- Insert Position of Drawing -->
  <xsl:template name="SetPosition">
    <xsl:variable name="InsertStartColumn">
      <xsl:call-template name="InsertStartColumn"/>
    </xsl:variable>
    <xsl:variable name="InsertStartColumnOffset">
      <xsl:call-template name="InsertStartColumnOffset"/>
    </xsl:variable>
    <xsl:variable name="InsertStartRow">
      <xsl:call-template name="InsertStartRow"/>
    </xsl:variable>
    <xsl:variable name="InsertStartRowOffset">
      <xsl:call-template name="InsertStartRowOffset"/>
    </xsl:variable>
    <xsl:variable name="InsertEndColumn">
      <xsl:call-template name="InsertEndColumn"/>
    </xsl:variable>
    <xsl:variable name="InsertEndColumnOffset">
      <xsl:call-template name="InsertEndColumnOffset"/>
    </xsl:variable>
    <xsl:variable name="InsertEndRow">
      <xsl:call-template name="InsertEndRow"/>
    </xsl:variable>
    <xsl:variable name="InsertEndRowOffset">
      <xsl:call-template name="InsertEndRowOffset"/>
    </xsl:variable>
    
    <xdr:from>
      <xdr:col>
        <xsl:choose>
          <xsl:when test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@draw:transform!= '' or parent::draw:frame/@draw:transform!= ''">
            <xsl:value-of select="$InsertStartColumn - ($InsertEndColumn - $InsertStartColumn) "/>    
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$InsertStartColumn"/>    
          </xsl:otherwise>
        </xsl:choose>
      </xdr:col>
      <xdr:colOff>
        <xsl:value-of select="$InsertStartColumnOffset"/>
      </xdr:colOff>
      <xdr:row>
        <xsl:choose>
          <xsl:when test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@draw:transform!= '' or parent::draw:frame/@draw:transform!= ''">
            <xsl:value-of select="$InsertStartRow - ($InsertEndRow - $InsertStartRow)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$InsertStartRow"/>
          </xsl:otherwise>
        </xsl:choose>        
      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$InsertStartColumnOffset"/>
      </xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>
        <xsl:choose>
          <xsl:when test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@draw:transform!= '' or parent::draw:frame/@draw:transform!= ''">
            <xsl:value-of select="$InsertStartColumn"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$InsertEndColumn"/>
          </xsl:otherwise>
        </xsl:choose>        
      </xdr:col>
      <xdr:colOff>
        <xsl:value-of select="$InsertEndColumnOffset"/>
      </xdr:colOff>
      <xdr:row>
        <xsl:choose>
        <xsl:when test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@draw:transform!= '' or parent::draw:frame/@draw:transform!= ''">
          <xsl:value-of select="$InsertStartRow"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$InsertEndRow"/>
        </xsl:otherwise>
      </xsl:choose>
      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$InsertEndRowOffset"/>
      </xdr:rowOff>
    </xdr:to>
  </xsl:template>

  <!-- Insert top left corner col number -->
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

  <!-- Insert top left corner row number -->
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

  <!-- Insert bottom right corner col number -->
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
  
  <!-- Insert bottom corner row number -->
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

  <!-- Horizontal offset of top left corner -->
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

  <!-- Vertical offset of top left corner -->
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
  
  <!-- Horizontal offset of  bottom right corner -->
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
  
  <!-- Vertical offset of  bottom right corner -->
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
  
  <!--image border width and line style-->
  <xsl:template name="InsertImageBorders">
    
    <xdr:spPr>
      
      <a:xfrm>
        <xsl:if test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@draw:transform!= '' or parent::draw:frame/@draw:transform!= ''">
          <xsl:attribute name="flipV">
            <xsl:text>1</xsl:text>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:mirror='horizontal'">
        <xsl:attribute name="flipH">
          <xsl:text>1</xsl:text>
        </xsl:attribute>
        </xsl:if>
      </a:xfrm>
      
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
      
      <xsl:variable name="BorderColor">
        <xsl:choose>
          <xsl:when test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@svg:stroke-color != ''">
            <xsl:value-of select="substring-after(key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@svg:stroke-color, '#')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>000000</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      
      <xsl:variable name="strokeWeight">        
         <xsl:call-template name="emu-measure">
           <xsl:with-param name="length" select="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@svg:stroke-width"/>          
          <xsl:with-param name="unit">emu</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      
      <a:ln> 
        <xsl:attribute name="w">
          <xsl:value-of select="$strokeWeight"/>
        </xsl:attribute>
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
            <xsl:value-of select="$BorderColor"/>
          </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </a:ln>
      
    
    </xdr:spPr>
    
  </xsl:template>

</xsl:stylesheet>

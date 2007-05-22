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
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  exclude-result-prefixes="w r xlink office draw text style">

  <xsl:template name="InsertPartRelationships">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

      <!-- Sheet relationship -->
      <xsl:for-each
        select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table">
        <Relationship Id="{generate-id(.)}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet">
          <xsl:variable name="NumberSheet">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:attribute name="Target">
            <xsl:value-of select="concat(concat('worksheets/sheet', $NumberSheet), '.xml')"/>
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>

      <!--  Static relationships -->
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        Target="styles.xml"/>
      <Relationship Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        Target="sharedStrings.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertWorksheetsRels">
    <xsl:param name="sheetNum"/>
    <xsl:param name="comment"/>
    <xsl:param name="picture"/>
    <xsl:param name="hyperlink"/>
    <xsl:param name="chart"/>

    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!-- comments.xml file -->
      <xsl:if test="$comment = 'true' ">
        <Relationship Id="{concat('rId',$sheetNum+1)}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
          Target="{concat('../comments',$sheetNum,'.xml')}"/>
      </xsl:if>

      <!-- vmlDrawing.vml file -->
      <xsl:if test="$comment = 'true' ">
        <Relationship Id="{concat('v_rId',$sheetNum)}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
          Target="{concat('../drawings/vmlDrawing',$sheetNum,'.vml')}"/>
      </xsl:if>

      <!-- drawing.xml file -->
      <xsl:if test="contains($chart,'true') or $picture = 'true'">
        <Relationship Id="{concat('d_rId',$sheetNum)}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
          Target="{concat('../drawings/drawing',$sheetNum,'.xml')}"/>
      </xsl:if>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertDrawingRels">
    <xsl:param name="sheetNum"/>

    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!-- chart rels -->
      <xsl:for-each
        select="descendant::draw:frame/draw:object[document(concat(translate(@xlink:href,'./',''),'/content.xml'))/office:document-content/office:body/office:chart]">
        <Relationship Id="{generate-id(.)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{concat('../charts/chart',$sheetNum,'_',position(),'.xml')}" />
      </xsl:for-each>

      <!-- TO DO: picture rels -->
      
      
      <xsl:for-each
        select="descendant::draw:frame/draw:image[contains(@xlink:href, 'Pictures')]">
        
    <xsl:variable name="imageName"
        select="substring-after(@xlink:href, 'Pictures/')"/>
      <pzip:copy pzip:source="{@xlink:href}" pzip:target="xl/media/{$imageName}"/>
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
        Id="{generate-id(.)}"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        Target="../media/{$imageName}"/>

        <!--Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
          Id="{generate-id(ancestor::draw:a)}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
          TargetMode="External">
          <xsl:attribute name="Target">
      <xsl:choose-->
              
              <!-- converting relative path -->
      <!--       <xsl:when test="starts-with(ancestor::draw:a/@xlink:href, '../')">
                <xsl:call-template name="HandlingSpaces">
                  <xsl:with-param name="path">
                    <xsl:value-of select="substring-after(ancestor::draw:a/@xlink:href,'../')"
                    />
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              
              <xsl:when test="starts-with(ancestor::draw:a/@xlink:href, '/')">
                <xsl:call-template name="HandlingSpaces">
                  <xsl:with-param name="path">
                    <xsl:value-of select="substring-after(ancestor::draw:a/@xlink:href,'/')"
                    />
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              
              <xsl:otherwise>
                <xsl:call-template name="HandlingSpaces">
                  <xsl:with-param name="path">
                    <xsl:value-of select="ancestor::draw:a/@xlink:href"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </Relationship> -->
        </xsl:for-each>
        </Relationships>
  </xsl:template>
</xsl:stylesheet>

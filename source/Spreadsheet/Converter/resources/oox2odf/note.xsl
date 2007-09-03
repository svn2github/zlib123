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
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  exclude-result-prefixes="e r draw">

  <xsl:import href="common.xsl"/>
  <!-- Get cell with note -->

  <xsl:template name="NoteCell">
    <xsl:param name="sheetNr"/>

    <xsl:variable name="targetFile">
      <xsl:value-of
        select="document(concat('xl/worksheets/_rels/sheet',$sheetNr,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'comments' )]/@Target"
      />
    </xsl:variable>

    <xsl:apply-templates
      select="document(concat('xl/',substring-after($targetFile,'../')))/e:comments"
      mode="note-cell"/>

  </xsl:template>

  <xsl:template match="e:comments" mode="note-cell">
    <xsl:apply-templates select="e:commentList/e:comment" mode="note-cell"/>
  </xsl:template>

  <xsl:template match="e:comment" mode="note-cell">

    <xsl:variable name="numCol">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="@ref"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="numRow">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="@ref"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat($numRow,':',$numCol,';')"/>
  </xsl:template>

  <!-- Get Row with Note -->
  <xsl:template name="NoteRow">
    <xsl:param name="NoteCell"/>
    <xsl:param name="Result"/>
    <xsl:choose>
      <xsl:when test="$NoteCell != ''">
        <xsl:call-template name="NoteRow">
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="substring-after($NoteCell, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="Result">
            <xsl:value-of select="concat($Result,  concat(substring-before($NoteCell, ':'), ';'))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$Result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Insert Note in This Cell -->
  <xsl:template name="InsertNoteInThisCell">
    <xsl:param name="rowNum"/>
    <xsl:param name="colNum"/>
    <xsl:param name="sheetNr"/>

    <xsl:variable name="thisCellCol">
      <xsl:call-template name="NumbersToChars">
        <xsl:with-param name="num">
          <xsl:value-of select="$colNum -1"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="thisCell">
      <xsl:value-of select="concat($thisCellCol,$rowNum)"/>
    </xsl:variable>

    <xsl:variable name="targetFile">
      <xsl:value-of
        select="document(concat('xl/worksheets/_rels/sheet',$sheetNr,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'comments' )]/@Target"
      />
    </xsl:variable>

    <xsl:variable name="fileNumber">
      <xsl:value-of select="substring-before(substring-after($targetFile,'comments'),'.xml')"/>
    </xsl:variable>

    <xsl:apply-templates
      select="document(concat('xl/comments',$fileNumber,'.xml'))/e:comments/e:commentList/e:comment[@ref=$thisCell]">
      <xsl:with-param name="number" select="$fileNumber"/>
      <xsl:with-param name="rowNum">
        <xsl:value-of select="$rowNum"/>
      </xsl:with-param>
      <xsl:with-param name="colNum">
        <xsl:value-of select="$colNum"/>
      </xsl:with-param>
    </xsl:apply-templates>

  </xsl:template>


  <!-- Insert Comment -->
  <xsl:template match="e:comment">
    <xsl:param name="rowNum"/>
    <xsl:param name="colNum"/>

    <!--@Description: adds a note -->
    <!--@context: none -->

    <xsl:param name="number"/>

    <!--(int) number of comments file -->
    <xsl:variable name="numberOfComment">
      <xsl:value-of select="count(preceding-sibling::e:comment)+1"/>
    </xsl:variable>

    <office:annotation>
     
      <xsl:variable name="widthInPt">
        <xsl:value-of select="substring-before(substring-after(document(concat('xl/drawings/vmlDrawing', $number, '.vml'))//v:shape[number($numberOfComment)]/@style, concat('width', ':')), ';')"/>
      </xsl:variable>
      
      <xsl:variable name="heightInPt">
        <xsl:value-of select="substring-before(substring-after(document(concat('xl/drawings/vmlDrawing', $number, '.vml'))//v:shape[number($numberOfComment)]/@style, concat('height', ':')), ';')"/>
      </xsl:variable>
            
      <xsl:attribute name="svg:width">
        <xsl:call-template name="pt2cm">
          <xsl:with-param name="pt">
            <xsl:value-of select="$widthInPt"/>
          </xsl:with-param>
        </xsl:call-template>  
      </xsl:attribute>
      
      <xsl:attribute name="svg:height">
        <xsl:call-template name="pt2cm">
          <xsl:with-param name="pt">
            <xsl:value-of select="$heightInPt"/>
          </xsl:with-param>
        </xsl:call-template>  
      </xsl:attribute>

      <xsl:for-each
        select="document(concat('xl/drawings/vmlDrawing',$number,'.vml'))/xml/v:shape[x:ClientData/x:Row = $rowNum - 1 and x:ClientData/x:Column = $colNum - 1]">

        <xsl:attribute name="office:display">
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'visibility'"/>
            <xsl:with-param name="shape" select="."/>
          </xsl:call-template>
        </xsl:attribute>

      </xsl:for-each>

      <xsl:apply-templates
        select="document(concat('xl/drawings/vmlDrawing',$number,'.vml'))/xml/v:shape[x:ClientData/x:Row = $rowNum - 1 and x:ClientData/x:Column = $colNum - 1]"
        mode="drawing">
        <xsl:with-param name="text" select="e:text"/>
      </xsl:apply-templates>

      <text:p text:style-name="{generate-id(e:text)}">
        <xsl:apply-templates select="e:text/e:r|e:text/e:t"/>
      </text:p>

    </office:annotation>

  </xsl:template>

  <xsl:template name="GetShapeProperty">
    <xsl:param name="shape"/>
    <xsl:param name="propertyName"/>

    <xsl:variable name="propertyValue">
      <xsl:choose>
        <xsl:when test="contains(substring-after($shape/@style,concat($propertyName,':')),';')">
          <xsl:value-of
            select="substring-before(substring-after($shape/@style,concat($propertyName,':')),';')"
          />
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="substring-after($shape/@style,concat($propertyName,';'))"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="CheckIfVisible">
      <xsl:choose>
        <xsl:when test="$propertyValue = 'visible'">
          <xsl:text>true</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>false</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="$CheckIfVisible"/>

  </xsl:template>

  <xsl:template name="InsertNoteStyles">
    <xsl:param name="sheetNr"/>

    <xsl:variable name="targetFile">
      <xsl:value-of
        select="document(concat('xl/worksheets/_rels/sheet',$sheetNr,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing' )]/@Target"
      />
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/', substring-after($targetFile, '/')))/xml/v:shape">
      <style:style style:name="{generate-id(.)}" style:family="graphic">
        <style:graphic-properties>

          <xsl:attribute name="fo:min-height">

            <xsl:variable name="height">
              <xsl:call-template name="ConvertPoints">
                <xsl:with-param name="length">
                  <xsl:value-of select="substring-before(substring-after(@style,'height:'),';')"/>
                </xsl:with-param>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>

            <xsl:choose>
              <xsl:when test="substring($height,1,1) = '.' ">
                <xsl:value-of select="concat('0',$height)"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$height"/>
              </xsl:otherwise>
            </xsl:choose>

          </xsl:attribute>

          <xsl:attribute name="draw:fill-color">
            <xsl:choose>
              <xsl:when test="@fillcolor">
                <xsl:call-template name="InsertColor"/>
                <xsl:value-of select="@fillcolor"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>

        </style:graphic-properties>
      </style:style>
    </xsl:for-each>

  </xsl:template>
  
 <!-- converts pt to centimeters -->
  <xsl:template name="pt2cm">
    <xsl:param name="pt"/>
    <xsl:variable name="ptValue">
      <xsl:if test="contains($pt, 'pt')">
        <xsl:value-of select="number(substring-before($pt, 'pt'))"/>
      </xsl:if>  
      <xsl:if test="not(contains($pt, 'pt'))">
        <xsl:value-of select="number($pt)"/>
      </xsl:if>
    </xsl:variable>
    
    <xsl:variable name="factor">
      <xsl:value-of select="127 div 3600"/>
    </xsl:variable>
    <xsl:value-of select="concat(string($ptValue * $factor), 'cm')"/>
  </xsl:template>

</xsl:stylesheet>

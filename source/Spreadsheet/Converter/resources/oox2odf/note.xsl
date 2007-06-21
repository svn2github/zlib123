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
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

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
    </xsl:apply-templates>

  </xsl:template>


  <!-- Insert Comment -->
  <xsl:template match="e:comment">

    <!--@Description: adds a note -->
    <!--@context: none -->

    <xsl:param name="number"/>

    <!--(int) number of comments file -->
    <xsl:variable name="numberOfComment">
      <xsl:value-of select="count(preceding-sibling::e:comment)+1"/>
    </xsl:variable>

    <office:annotation>
      <xsl:apply-templates
        select="document(concat('xl/drawings/vmlDrawing',$number,'.vml'))/xml/v:shape[position()=$numberOfComment]"
        mode="drawing">
        <xsl:with-param name="text" select="e:text"/>
      </xsl:apply-templates>
      <text:p text:style-name="{generate-id(e:text)}">
        <xsl:apply-templates select="e:text/e:r"/>
      </text:p>
    </office:annotation>
  </xsl:template>

</xsl:stylesheet>

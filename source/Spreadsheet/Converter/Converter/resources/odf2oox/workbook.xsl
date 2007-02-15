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
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"  
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  	xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    
  <xsl:import href="worksheets.xsl"/>
  <xsl:import href="sharedStrings.xsl"/>
  
  
  <!-- main workbook template-->
    <xsl:template name="InsertWorkbook">    
      <xsl:apply-templates select="document('content.xml')/office:document-content"/>
    </xsl:template>
    
  <!-- workbook body template -->
    <xsl:template match="office:body">
    <workbook>
      <xsl:apply-templates select="office:spreadsheet"/>
    </workbook>
    </xsl:template>
  
  <!-- insert references to all sheets -->
  <xsl:template match="office:spreadsheet">
      <sheets>
        <xsl:for-each select="table:table">
          <sheet name="{@table:name}" sheetId="{count(preceding-sibling::table:table)+1}" r:id="{generate-id(.)}"/>
        </xsl:for-each>
      </sheets>
  </xsl:template>
  
  <!-- insert all sheets -->
  <xsl:template name="InsertSheets">
    <xsl:apply-templates select="document('content.xml')/office:document-content" mode="sheet"/>
  </xsl:template>
  
</xsl:stylesheet>
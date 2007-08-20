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
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    exclude-result-prefixes="#default w r office style table">
    
    <xsl:template name="revisionHeaders">
        <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:tracked-changes">         
            <pzip:entry pzip:target="xl/revisions/revisionHeaders.xml">
                <xsl:call-template name="revisionHeaderProperties"/>
            </pzip:entry>        
        </xsl:for-each>
    </xsl:template>
    
    <!-- Create Revisions Files -->
    <xsl:template name="CreateRevisionFiles">
        <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:tracked-changes">                
            <xsl:apply-templates select="node()[1][name()='table:cell-content-change']" mode="revisionFiles"/>
        </xsl:for-each>
    </xsl:template>
    
    <xsl:template match="table:cell-content-change" mode="revisionFiles">
        
        <pzip:entry pzip:target="{concat('xl/revisions/revisionLog', generate-id(), '.xml')}">
            <xsl:call-template name="InsertChangeTrackingProperties"/>
        </pzip:entry>
        
        <xsl:if test="following-sibling::node()[1][name()='table:cell-content-change']">
            <xsl:apply-templates select="following-sibling::node()[1][name()='table:cell-content-change']" mode="revisionFiles"/>
        </xsl:if>
        
    </xsl:template>
    
    <xsl:template name="userName">        
        <pzip:entry pzip:target="xl/revisions/userNames.xml">
            <users xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="0"/>
        </pzip:entry>
    </xsl:template>
    
    <!-- Insert Change Tracking Properties -->
    <xsl:template name="InsertChangeTrackingProperties">
        <revisions xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <rcc rId="1" sId="1">
                <oc r="H10" t="inlineStr">
                    <is>
                        <t>trkst</t>
                    </is>
                </oc>
                <nc r="H10" t="inlineStr">
                    <is>
                        <t>tekst1</t>
                    </is>
                </nc>
            </rcc>
        </revisions>
    </xsl:template>
    
  
    
</xsl:stylesheet>
<?xml version="1.0" encoding="UTF-8"?>
<!--
    * Copyright (c) 2006, Clever Age
    * All rights reserved.
    * 
    * Redistribution and use in source and binary forms, with or without
    * modification, are permitted provided that <office:spreadsheet the following conditions are met:
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
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
    xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
    xmlns:math="http://www.w3.org/1998/Math/MathML"
    xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
    xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
    xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer"
    xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events"
    xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.0"
    xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">
    
    <xsl:template name="InsertChangeTracking">
        
        <!-- Search Revision Log Files -->
        <xsl:variable name="revisionLogFiles">
            <xsl:for-each
                select="document('xl/_rels/workbook.xml.rels')//node()[name()='Relationship']">
                <xsl:if test="./@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders'">
                <xsl:value-of select="./@Target"/>
                </xsl:if>
            </xsl:for-each>
        </xsl:variable>
        
        <!-- Open Files with Change Tracking -->
        
        <xsl:for-each select="document(concat('xl/', $revisionLogFiles))">
            <table:tracked-changes>
                <xsl:call-template name="SearchChangeTracking">
                    <xsl:with-param name="revisionLogFiles">
                        <xsl:value-of select="$revisionLogFiles"/>
                    </xsl:with-param>
                </xsl:call-template>
            </table:tracked-changes>
        </xsl:for-each>
        
    </xsl:template>
    
   <!-- Insert Change Tracking -->
    
    <xsl:template name="SearchChangeTracking">
        <xsl:param name="revisionLogFiles"/>
        
        
        <xsl:for-each select="e:headers">
           <xsl:apply-templates select="e:header[1]">
               <xsl:with-param name="revisionLogFiles">
                   <xsl:value-of select="$revisionLogFiles"/>
               </xsl:with-param>
           </xsl:apply-templates>
        </xsl:for-each>
        
    </xsl:template>
    
    <!-- Insert Change Tracking Properties -->
    
    <xsl:template match="e:header">
        <xsl:param name="revisionLogFiles"/>
        <xsl:param name="positionHeader" select="1"/>
        
        <xsl:variable name="userName">
            <xsl:value-of select="@userName"/>
        </xsl:variable>
        
        <xsl:variable name="dateTime">
            <xsl:value-of select="@dateTime"/>
        </xsl:variable>
        
        <xsl:variable name="id">
           <xsl:value-of select="@r:id"/>
        </xsl:variable>        
        
            <xsl:variable name="Target">
            <xsl:for-each
                select="document(concat('xl/', substring-before($revisionLogFiles, '/'), '/_rels/', substring-after($revisionLogFiles, '/'), '.rels'))//node()[name()='Relationship']">
                <xsl:if test="./@Id=$id">
                    <xsl:value-of select="concat('xl/', substring-before($revisionLogFiles, '/'), '/', ./@Target)"/>
                </xsl:if>
            </xsl:for-each>
            </xsl:variable>
            
            <xsl:for-each select="document($Target)/e:revisions">
                <xsl:if test="e:rcc">
                <xsl:apply-templates select="e:rcc">
                    <xsl:with-param name="dateTime">
                        <xsl:value-of select="$dateTime"/>
                    </xsl:with-param>
                    <xsl:with-param name="userName">
                        <xsl:value-of select="$userName"/>
                    </xsl:with-param>
                    <xsl:with-param name="positionHeader">
                        <xsl:value-of select="$positionHeader"/>
                    </xsl:with-param>
                    </xsl:apply-templates>
                </xsl:if>
            </xsl:for-each>
         
        <xsl:if test="following-sibling::e:header">
            <xsl:apply-templates select="following-sibling::e:header[1]">
                <xsl:with-param name="revisionLogFiles">
                    <xsl:value-of select="$revisionLogFiles"/>
                </xsl:with-param>
                <xsl:with-param name="positionHeader">
                    <xsl:value-of select="$positionHeader + 1"/>
                </xsl:with-param>
            </xsl:apply-templates>
        </xsl:if>
        
    </xsl:template>
    
    <xsl:template match="e:rcc">
        <xsl:param name="dateTime"/>
        <xsl:param name="userName"/>
        <xsl:param name="positionHeader"/>
        
        <table:cell-content-change>
            
            <xsl:attribute name="table:id">
                <xsl:value-of select="concat('ct', $positionHeader, position())"/>
            </xsl:attribute>
            
            <xsl:variable name="colNum">
                <xsl:call-template name="GetColNum">
                    <xsl:with-param name="cell">
                        <xsl:choose>
                            <xsl:when test="e:oc/@r">
                                <xsl:value-of select="e:oc/@r"/>
                            </xsl:when>
                            <xsl:when test="e:nc/@r">
                                <xsl:value-of select="e:nc/@r"/>
                            </xsl:when>
                        </xsl:choose>
                    </xsl:with-param>
                </xsl:call-template>
            </xsl:variable>
            
            <xsl:variable name="rowNum">
                <xsl:call-template name="GetRowNum">
                    <xsl:with-param name="cell">
                        <xsl:choose>
                            <xsl:when test="e:oc/@r">
                                <xsl:value-of select="e:oc/@r"/>
                            </xsl:when>
                            <xsl:when test="e:nc/@r">
                                <xsl:value-of select="e:nc/@r"/>
                            </xsl:when>
                        </xsl:choose>
                    </xsl:with-param>
                </xsl:call-template>
            </xsl:variable>
            
            <table:cell-address>
                <xsl:attribute name="table:table">
                    <xsl:value-of select="@sId - 1"/>
                </xsl:attribute>
                <xsl:attribute name="table:row">
                    <xsl:value-of select="$rowNum - 1"/>
                </xsl:attribute>
                <xsl:attribute name="table:column">
                    <xsl:value-of select="$colNum - 1"/>
                </xsl:attribute>
            </table:cell-address>
            
            
            <office:change-info>
                <dc:creator>
                    <xsl:value-of select="$userName"/>
                </dc:creator>
                <dc:date>
                    <xsl:value-of select="$dateTime"/>
                </dc:date>
            </office:change-info>
        
        <table:previous>
            <table:change-track-table-cell office:value-type="string">
                <xsl:if test="e:oc/e:is/e:t">
                <text:p>
                    <xsl:value-of select="e:oc/e:is/e:t"/>
                </text:p>
                </xsl:if>
            </table:change-track-table-cell>
        </table:previous>
            
        </table:cell-content-change>
        
    </xsl:template>
    
   
    
    
    </xsl:stylesheet>
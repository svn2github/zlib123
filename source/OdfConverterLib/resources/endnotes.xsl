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
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:zip="urn:cleverage:xmlns:zip"
    exclude-result-prefixes="text style office xlink draw zip">
    
    <xsl:key name="endnotes" match="text:note[@text:note-class='endnote']" use="''"/>
    
    <xsl:template name="endnotes">
        <w:endnotes>
            
            <!-- special endnotes -->
            <w:endnote w:type="separator" w:id="0">
                <w:p>
                    <!-- there are no styles for endnotes separator in oo -->
                    <w:pPr>
                        <w:spacing w:after="0" w:line="240"
                            w:lineRule="auto"/>
                    </w:pPr>
                    <w:r>
                        <w:separator/>
                    </w:r>
                </w:p>
            </w:endnote>
            <w:endnote w:type="continuationSeparator" w:id="1">
                <w:p>
                      <w:pPr>
                        <w:spacing w:after="0" w:line="240"
                            w:lineRule="auto"/>
                    </w:pPr>
                    <w:r>
                        <w:continuationSeparator/>
                    </w:r>
                </w:p>
            </w:endnote>
            
            <!-- normal endnotes -->
            
            <xsl:for-each select="document('content.xml')">
                <xsl:for-each select="key('endnotes', '')">
                    <w:endnote w:type="normal" w:id="{position() + 1}">
                        <xsl:if test="text:note-citation/@text:label">
                            <xsl:attribute name="w:suppressRef"
                                >1</xsl:attribute>
                        </xsl:if>
                        <xsl:apply-templates select="text:note-body"/>
                    </w:endnote>
                </xsl:for-each>
            </xsl:for-each>
            
        </w:endnotes>
    </xsl:template>
    
    <!-- endotes configuration -->
    <xsl:template name="endnotes-configuration">
        <xsl:variable name="config"
            select="document('styles.xml')/office:document-styles/office:styles/text:notes-configuration[@text:note-class='endnote']"/>
        <w:endnotePr>
            
            <!-- TODO endnotes for sections -->
            <w:pos>
                <xsl:attribute name="w:val">docEnd</xsl:attribute>
            </w:pos>
            
            <w:numFmt>
                <xsl:attribute name="w:val">
                    <xsl:call-template name="num-format">
                        <xsl:with-param name="format"
                            select="$config/@style:num-format"/>
                    </xsl:call-template>
                </xsl:attribute>
            </w:numFmt>
            
            <xsl:if test="$config/@text:start-value">
                <w:numStart w:val="{$config/@text:start-value}"/>
            </xsl:if>
            
            <w:numRestart>
                <xsl:attribute name="w:val">
                    <xsl:choose>
                        <xsl:when
                            test="$config/@text:start-numbering-at = 'page' "
                            >eachPage</xsl:when>
                        <xsl:when
                            test="$config/@text:start-numbering-at = 'chapter' "
                            >eachSect</xsl:when>
                        <xsl:otherwise>continuous</xsl:otherwise>
                    </xsl:choose>
                </xsl:attribute>
            </w:numRestart>
            
            <w:endnote w:id="0"/>
            <w:endnote w:id="1"/>
            
        </w:endnotePr>
    </xsl:template>
    
    <!--TODO endnotes relationships-->
    
</xsl:stylesheet>

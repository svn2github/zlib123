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
    exclude-result-prefixes="text">
    
        <xsl:key name="footnotes" match="text:note" use="''"/>
        
        <xsl:template name="footnotes">
                <w:footnotes>
                        
                        <!-- special footnotes -->
                        <w:footnote w:type="separator" w:id="0">
                                <w:p>
                                        <!-- TODO :  extract properties from document('styles.xml')/office:automatic-styles/style:page-layout/style:footnote-sep -->
                                        <w:pPr>
                                                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                                        </w:pPr>
                                        <w:r>
                                                <w:separator/>
                                        </w:r>
                                </w:p>
                        </w:footnote>
                        <w:footnote w:type="continuationSeparator" w:id="1">
                                <w:p>
                                        <!-- TODO :  extract properties from document('styles.xml')/office:automatic-styles/style:page-layout/style:footnote-sep -->
                                        <w:pPr>
                                                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                                        </w:pPr>
                                        <w:r>        
                                                <w:continuationSeparator/>
                                        </w:r>
                                </w:p>
                        </w:footnote>
                        
                        <!-- normal footnotes -->
                        
                        <!-- absurd hack for changing the context -->
                        <xsl:for-each select="document('content.xml')">
                                <!-- iterating over the footnotes -->
                                    <xsl:for-each select="key('footnotes', '')">
                                            <w:footnote w:type="normal" w:id="{position() + 1}">
                                                    <w:p>
                                                            <w:pPr>
                                                                    <w:pStyle w:val="{concat(@text:note-class, 'Text')}"/>
                                                            </w:pPr>
                                                            <w:r>
                                                                    <w:rPr>
                                                                            <w:rStyle w:val="{concat(@text:note-class, 'Reference')}"/>
                                                                    </w:rPr>
                                                                    <w:footnoteRef/>
                                                             </w:r>
                                                         <xsl:apply-templates select="text:note-body" mode="paragraph"/>   
                                                    </w:p>
                                            </w:footnote>
                                    </xsl:for-each>
                        </xsl:for-each>
                        
                </w:footnotes>
        </xsl:template>
 
</xsl:stylesheet>

  




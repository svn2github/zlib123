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
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/3/wordprocessingDrawing"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    exclude-result-prefixes="xlink draw svg fo">


    <xsl:key name="images" match="draw:frame" use="'const'"/>

    <!-- Get the position of an element in the draw:frame group -->
    <xsl:template name="GetPosition">
        <xsl:param name="node"/>
        <xsl:variable name="positionInGroup">
            <xsl:for-each select="key('images', 'const')">
                <xsl:if test="generate-id($node) = generate-id(.)">
                    <xsl:value-of select="position()"/>
                </xsl:if>
            </xsl:for-each>
        </xsl:variable>
        <xsl:value-of select="$positionInGroup"/>
    </xsl:template>

    <!-- TODO: manage ole-objects -->
    <xsl:template match="draw:frame[not(./draw:object-ole) and starts-with(./draw:image/@xlink:href, 'Pictures/')]" mode="paragraph">
        <xsl:variable name="intId">
            <xsl:call-template name="GetPosition">
                <xsl:with-param name="node" select="."/>
            </xsl:call-template>
        </xsl:variable>

        <xsl:variable name="cx">
            <xsl:choose>
                <xsl:when test="@svg:width">
                    <xsl:call-template name="emu-measure">
                        <xsl:with-param name="length" select="@svg:width"/>
                    </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:call-template name="emu-measure">
                        <xsl:with-param name="length" select="draw:text-box/@fo:min-width"/>
                    </xsl:call-template>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>

        <xsl:variable name="cy">
            <xsl:choose>
                <xsl:when test="@svg:height">
                    <xsl:call-template name="emu-measure">
                        <xsl:with-param name="length" select="@svg:height"/>
                    </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:call-template name="emu-measure">
                        <xsl:with-param name="length" select="draw:text-box/@fo:min-height"/>
                    </xsl:call-template>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        
        <xsl:variable name="ox">
            <xsl:choose>
                <xsl:when test="@svg:x">
                    <xsl:call-template name="emu-measure">
                        <xsl:with-param name="length" select="@svg:x"/>
                    </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="0"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        
        <xsl:variable name="oy">
            <xsl:choose>
                <xsl:when test="@svg:y">
                    <xsl:call-template name="emu-measure">
                        <xsl:with-param name="length" select="@svg:y"/>
                    </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="0"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        
        <xsl:variable name="sName" select="@draw:style-name"/>
        
        <xsl:variable name="style" select="//style:style[@style:name = $sName]/style:graphic-properties"/>
        
        <xsl:variable name="posH" select="$style/@style:horizontal-rel"/>
        
        <xsl:variable name="wrap" select="$style/@style:wrap"/>

        <xsl:choose>
            <xsl:when test="@svg:x or @svg:y">
                <w:r>
                    <w:drawing>
                        <wp:anchor simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0"
                            layoutInCell="1" allowOverlap="1">
                            <wp:simplePos x="0" y="0"/>
                            <wp:positionH>
                                <xsl:attribute name="relativeFrom">
                                    <xsl:choose>
                                        <xsl:when test="@text:anchor-type = 'page'">page</xsl:when>
                                        <xsl:otherwise>column</xsl:otherwise>
                                    </xsl:choose>
                                </xsl:attribute>
                                <wp:posOffset>
                                    <xsl:value-of select="$ox"/>
                                </wp:posOffset>
                            </wp:positionH>
                            <wp:positionV>
                                <xsl:attribute name="relativeFrom">
                                    <xsl:choose>
                                        <xsl:when test="@text:anchor-type = 'page'">page</xsl:when>
                                        <xsl:otherwise>paragraph</xsl:otherwise>
                                    </xsl:choose>
                                </xsl:attribute>
                                <wp:posOffset>
                                    <xsl:value-of select="$oy"/>
                                </wp:posOffset>
                            </wp:positionV>
                            <wp:extent cx="{$cx}" cy="{$cy}"/>
                            <wp:effectExtent l="0" t="0" r="0" b="0"/>
                            <xsl:choose>
                                <xsl:when test="$wrap = 'parallel' or $wrap ='left' or $wrap = 'right' or $wrap ='dynamic'">
                                    <wp:wrapSquare>
                                        <xsl:attribute name="wrapText">
                                            <xsl:choose>
                                                <xsl:when test="$wrap = 'parallel'">bothSides</xsl:when>
                                                <xsl:when test="$wrap = 'dynamic'">largest</xsl:when>
                                                <xsl:otherwise>
                                                    <xsl:value-of select="$wrap"/>
                                                </xsl:otherwise>
                                            </xsl:choose>
                                        </xsl:attribute>
                                        <xsl:attribute name="distB">
                                            <xsl:choose>
                                                <xsl:when test="$style/@fo:margin-bottom">
                                                    <xsl:call-template name="emu-measure">
                                                        <xsl:with-param name="length" select="$style/@fo:margin-bottom"/>
                                                    </xsl:call-template>    
                                                </xsl:when>
                                                <xsl:otherwise><xsl:value-of select="0"/></xsl:otherwise>
                                            </xsl:choose>    
                                         </xsl:attribute>    
                                        <xsl:attribute name="distT">
                                            <xsl:choose>
                                                <xsl:when test="$style/@fo:margin-top">
                                                    <xsl:call-template name="emu-measure">
                                                        <xsl:with-param name="length" select="$style/@fo:margin-top"/>
                                                    </xsl:call-template>    
                                                </xsl:when>
                                                <xsl:otherwise><xsl:value-of select="0"/></xsl:otherwise>
                                            </xsl:choose>
                                        </xsl:attribute>
                                        <xsl:attribute name="distL">
                                            <xsl:choose>
                                                <xsl:when test="$style/@fo:margin-left">
                                                    <xsl:call-template name="emu-measure">
                                                        <xsl:with-param name="length" select="$style/@fo:margin-left"/>
                                                    </xsl:call-template>
                                                </xsl:when>
                                                <xsl:otherwise><xsl:value-of select="0"/></xsl:otherwise>
                                            </xsl:choose>
                                        </xsl:attribute>
                                        <xsl:attribute name="distR">
                                            <xsl:choose>
                                                <xsl:when test="$style/@fo:margin-right">
                                                    <xsl:call-template name="emu-measure">
                                                        <xsl:with-param name="length" select="$style/@fo:margin-right"/>
                                                    </xsl:call-template>
                                                </xsl:when>
                                                <xsl:otherwise><xsl:value-of select="0"/></xsl:otherwise>
                                            </xsl:choose>
                                        </xsl:attribute>
                                    </wp:wrapSquare>
                                </xsl:when>
                                <xsl:when test="$wrap = 'run-through'">
                                    <wp:wrapNone/>
                                </xsl:when>
                                <xsl:when test="$wrap = 'none'">
                                    <wp:wrapTopAndBottom>
                                        <xsl:attribute name="distB">
                                            <xsl:choose>
                                                <xsl:when test="$style/@fo:margin-bottom">
                                                    <xsl:call-template name="emu-measure">
                                                        <xsl:with-param name="length" select="$style/@fo:margin-bottom"/>
                                                    </xsl:call-template>
                                                </xsl:when>
                                                <xsl:otherwise>
                                                    <xsl:value-of select="0"/>
                                                </xsl:otherwise>
                                            </xsl:choose>
                                         </xsl:attribute>
                                         <xsl:attribute name="distT">
                                             <xsl:choose>
                                                 <xsl:when test="$style/@fo:margin-top">
                                                     <xsl:call-template name="emu-measure">
                                                        <xsl:with-param name="length" select="$style/@fo:margin-top"/>
                                                     </xsl:call-template>
                                                 </xsl:when>
                                                 <xsl:otherwise>
                                                     <xsl:value-of select="0"/>
                                                 </xsl:otherwise>
                                             </xsl:choose>
                                            </xsl:attribute>
                                    </wp:wrapTopAndBottom>
                                </xsl:when>
                                <xsl:otherwise>
                                    <wp:wrapNone/>
                                </xsl:otherwise>
                            </xsl:choose>
                            <wp:docPr name="{@draw:name}" id="{$intId}"/>
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks
                                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/3/main"
                                    noChangeAspect="1"/>
                            </wp:cNvGraphicFramePr>
                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/3/main">
                                <a:graphicData
                                    uri="http://schemas.openxmlformats.org/drawingml/2006/3/picture">
                                    <pic:pic
                                        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/3/picture">
                                        <pic:nvPicPr>
                                            <pic:cNvPr name="{@draw:name}" id="{$intId}"/>
                                            <pic:cNvPicPr>
                                                <a:picLocks noChangeAspect="1"/>
                                            </pic:cNvPicPr>
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                           <a:blip r:embed="{generate-id(draw:image)}"/>
                                            <a:stretch>
                                                <a:fillRect/>
                                            </a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm>
                                                <a:off x="0" y="0"/>
                                                <a:ext cx="{$cx}" cy="{$cy}"/>
                                            </a:xfrm>
                                            <a:prstGeom prst="rect">
                                                <a:avLst/>
                                            </a:prstGeom>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:anchor>
                    </w:drawing>
                </w:r>
            </xsl:when>
            <xsl:otherwise>
                <w:r>
                        <w:drawing>
                            <wp:inline>
                                <!-- width and heigth -->
                                <wp:extent cx="{$cx}" cy="{$cy}"/>
                                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                                <wp:docPr name="{@draw:name}" id="{$intId}"/>
                                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/3/main">
                                    <a:graphicData
                                        uri="http://schemas.openxmlformats.org/drawingml/2006/3/picture">
                                        <pic:pic
                                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/3/picture">
                                            <!-- non visual drawing properties -->
                                            <pic:nvPicPr>
                                                <pic:cNvPr name="{@draw:name}" id="{$intId}"/>
                                                <pic:cNvPicPr>
                                                    <!-- TODO : implement  cNvPicPr -->
                                                    <a:picLocks noChangeAspect="1"/>
                                                </pic:cNvPicPr>
                                            </pic:nvPicPr>
                                            <pic:blipFill>
                                                <a:blip r:embed="{generate-id(draw:image)}"/>
                                                <a:stretch>
                                                    <!-- TODO -->
                                                    <a:fillRect/>
                                                </a:stretch>
                                            </pic:blipFill>
                                            <pic:spPr>
                                                <a:xfrm>
                                                    <a:off x="0" y="0"/>
                                                    <!-- TODO -->
                                                    <a:ext cx="{$cx}" cy="{$cy}"/>
                                                </a:xfrm>
                                                <!-- TODO -->
                                                <a:prstGeom prst="rect">
                                                    <a:avLst/>
                                                </a:prstGeom>
                                            </pic:spPr>
                                        </pic:pic>
                                    </a:graphicData>
                                </a:graphic>
                            </wp:inline>
                        </w:drawing>
                    </w:r>
            </xsl:otherwise>
        </xsl:choose>
        
    </xsl:template>


</xsl:stylesheet>

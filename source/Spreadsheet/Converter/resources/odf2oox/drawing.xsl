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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" exclude-result-prefixes="a">

  <xsl:import href="cell.xsl"/>
  <xsl:import href="common.xsl"/>

  <!-- Insert Drawing (picture, chart)  -->
  <xsl:template name="InsertDrawing">



    <xdr:wsDr>
      <!--Insert Chart -->
      <xsl:for-each select="descendant::draw:frame">

        <xsl:variable name="chart">
          <xsl:for-each select="descendant::draw:object">
            <xsl:choose>
              <xsl:when test="not(document(concat(translate(@xlink:href,'./',''),'/settings.xml')))">
                <xsl:for-each
                  select="document(concat(translate(@xlink:href,'./',''),'/content.xml'))">
                  <xsl:choose>
                    <xsl:when test="office:document-content/office:body/office:chart">
                      <xsl:text>true</xsl:text>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:text>false</xsl:text>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>false</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:variable>

        <xsl:choose>
          <!-- insert chart -->
          <xsl:when test="contains($chart, 'true')">
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
          </xsl:when>

          <!-- insert picture -->
          <xsl:when
            test="draw:image[not(name(parent::node()/parent::node()) = 'draw:g' ) and not(parent::node()/draw:object)]">
            <xdr:twoCellAnchor>
              <xsl:call-template name="SetPosition"/>
              <xdr:pic>
                <xdr:nvPicPr>
                  <xdr:cNvPr>
                    <xsl:attribute name="id">
                      <xsl:value-of select="position()"/>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:value-of select="@draw:name"/>
                    </xsl:attribute>
                    <xsl:attribute name="descr">
                      <xsl:value-of select="@draw:name"/>
                    </xsl:attribute>
                  </xdr:cNvPr>
                  <xdr:cNvPicPr>
                    <a:picLocks noChangeAspect="1"/>
                  </xdr:cNvPicPr>
                </xdr:nvPicPr>

                <xdr:blipFill>
                  <a:blip
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <xsl:choose>
                      <!-- embeded picture -->
                      <xsl:when test="starts-with(draw:image/@xlink:href, 'Pictures/')">
                        <xsl:attribute name="r:embed">
                          <xsl:value-of select="generate-id()"/>
                        </xsl:attribute>
                      </xsl:when>
                      <!-- linked picture -->
                      <xsl:otherwise>
                        <xsl:attribute name="r:link">
                          <xsl:value-of select="generate-id()"/>
                        </xsl:attribute>
                      </xsl:otherwise>
                    </xsl:choose>
                  </a:blip>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </xdr:blipFill>

                <xsl:call-template name="InsertFrameProperties"/>

              </xdr:pic>
              <xdr:clientData/>
            </xdr:twoCellAnchor>
          </xsl:when>

          <!-- insert text-box -->
          <xsl:when test="draw:text-box">
            <xdr:twoCellAnchor>

              <xsl:call-template name="SetPosition"/>

              <xdr:sp macro="" textlink="">
                <xdr:nvSpPr>
                  <xdr:cNvPr id="{position()}" name="{concat('TextBox ',position())}"/>
                  <xdr:cNvSpPr txBox="1"/>
                </xdr:nvSpPr>

                <xsl:call-template name="InsertFrameProperties"/>

                <xdr:txBody>

                  <xsl:call-template name="InsertTextBoxProperties"/>

                  <a:lstStyle/>

                  <!-- insert text -->
                  <xsl:apply-templates mode="text-box"/>

                </xdr:txBody>

              </xdr:sp>
              <xdr:clientData/>
            </xdr:twoCellAnchor>

          </xsl:when>
        </xsl:choose>
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
          <xsl:when
            test="key('style', @draw:style-name)/style:graphic-properties/@draw:transform!= '' or @draw:transform!= ''">
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
          <xsl:when
            test="key('style', @draw:style-name)/style:graphic-properties/@draw:transform!= '' or @draw:transform!= ''">
            <xsl:value-of select="$InsertStartRow - ($InsertEndRow - $InsertStartRow)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$InsertStartRow"/>
          </xsl:otherwise>
        </xsl:choose>
      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$InsertStartRowOffset"/>
      </xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>

        <!-- To do  (check roundtrip Diagrammtypen.xlsx)-->
        <xsl:choose>
          <xsl:when
            test="key('style', @draw:style-name)/style:graphic-properties/@draw:transform!= '' or @draw:transform!= '' and $InsertStartColumn != '' and $InsertStartColumn != 'NaN'">
            <xsl:value-of select="$InsertStartColumn"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="$InsertEndColumn != '' and $InsertEndColumn != 'NaN'">
                <xsl:value-of select="$InsertEndColumn"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$InsertStartColumn"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>

      </xdr:col>
      <xdr:colOff>
        <xsl:value-of select="$InsertEndColumnOffset"/>
      </xdr:colOff>
      <xdr:row>
        <xsl:choose>
          <xsl:when
            test="key('style', @draw:style-name)/style:graphic-properties/@draw:transform!= '' or @draw:transform!= ''">
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
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <xsl:for-each select="parent::node()">
          <xsl:variable name="position">
            <xsl:value-of
              select="count(preceding-sibling::table:table-cell) + count(preceding-sibling::table:covered-table-cell) + 1"
            />
          </xsl:variable>
          <xsl:variable name="number">
            <xsl:for-each
              select="parent::node()/child::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ][1]">
              <xsl:call-template name="GetColNumber">
                <xsl:with-param name="position" select="$position"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:variable>
          <xsl:value-of select="$number - 1"/>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>1</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Insert top left corner row number -->
  <xsl:template name="InsertStartRow">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <!-- get parent table:table-row id -->
        <xsl:variable name="rowId">
          <xsl:value-of select="generate-id(ancestor::table:table-row)"/>
        </xsl:variable>
        <!-- go to first table:table-row-->
        <xsl:for-each select="ancestor::table:table/descendant::table:table-row[1]">
          <xsl:variable name="number">
            <xsl:call-template name="GetRowNumber">
              <xsl:with-param name="rowId" select="$rowId"/>
              <xsl:with-param name="tableId" select="generate-id(ancestor::table:table)"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="$number - 1"/>
        </xsl:for-each>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>31</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Insert bottom right corner col number -->
  <xsl:template name="InsertEndColumn">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <xsl:variable name="number">
          <xsl:call-template name="GetColNum">
            <xsl:with-param name="cell" select="substring-after(@table:end-cell-address,'.')"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$number - 1"/>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>5</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Insert bottom corner row number -->
  <xsl:template name="InsertEndRow">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <xsl:variable name="number">
          <xsl:call-template name="GetRowNum">
            <xsl:with-param name="cell" select="substring-after(@table:end-cell-address,'.')"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$number - 1"/>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>46</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Horizontal offset of top left corner -->
  <xsl:template name="InsertStartColumnOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="@svg:x"/>
        </xsl:call-template>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>714375</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Vertical offset of top left corner -->
  <xsl:template name="InsertStartRowOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="@svg:y"/>
        </xsl:call-template>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>104775</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Horizontal offset of  bottom right corner -->
  <xsl:template name="InsertEndColumnOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="@ table:end-x"/>
        </xsl:call-template>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>447675</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Vertical offset of  bottom right corner -->
  <xsl:template name="InsertEndRowOffset">
    <xsl:choose>
      <!-- when anchor is to cell -->
      <xsl:when
        test="parent::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ]">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="@ table:end-y"/>
        </xsl:call-template>
      </xsl:when>
      <!-- when anchor is to page -->
      <xsl:when test="parent::node()[name() = 'table:shapes']">
        <xsl:text>104775</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--image border width and line style-->
  <xsl:template name="InsertFrameProperties">

    <xdr:spPr>

      <a:xfrm>
        <xsl:if
          test="key('style', @draw:style-name)/style:graphic-properties/@draw:transform!= '' or @draw:transform!= ''">
          <xsl:attribute name="flipV">
            <xsl:text>1</xsl:text>
          </xsl:attribute>
        </xsl:if>
        <xsl:if
          test="key('style', @draw:style-name)/style:graphic-properties/@style:mirror='horizontal'">
          <xsl:attribute name="flipH">
            <xsl:text>1</xsl:text>
          </xsl:attribute>
        </xsl:if>
      </a:xfrm>

      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>

      <!--xsl:if test="key('style', @draw:style-name)/style:graphic-properties/@draw:stroke != 'none' "-->

      <xsl:for-each select="key('style', @draw:style-name)/style:graphic-properties">
        <xsl:call-template name="InsertDrawingFill"/>
        <xsl:call-template name="InsertDrawingBorder"/>
      </xsl:for-each>

      <!--/xsl:if-->

    </xdr:spPr>

  </xsl:template>

  <xsl:template match="text:p" mode="text-box">
    <xsl:variable name="tekst">
      <xsl:value-of select="."/>
    </xsl:variable>

    <a:p>
      <xsl:call-template name="InsertParagraphProperties"/>

      <xsl:choose>
        <!-- empty line -->
        <xsl:when test="$tekst = '' and not(text:s) and not(text:span)">
          <xsl:call-template name="InsertEndPara"/>
        </xsl:when>

        <!-- text:p without text:span -->
        <xsl:when test="not(text:span)">
          <a:r>
            <a:rPr lang="en-EN" sz="1200">
              <xsl:for-each select="key('style',@text:style-name)/style:text-properties">
                <xsl:call-template name="InsertRunProperties"/>
              </xsl:for-each>

              <!-- default font face -->
              <xsl:if test="not(key('style',@text:style-name)/style:text-properties)">
                <a:latin typeface="Times New Roman" charset="0"/>
              </xsl:if>

            </a:rPr>
            <a:t>
              <xsl:apply-templates mode="text-box"/>
            </a:t>
          </a:r>
        </xsl:when>

        <xsl:otherwise>
          <xsl:apply-templates mode="text-box"/>
        </xsl:otherwise>
      </xsl:choose>
    </a:p>
  </xsl:template>

  <xsl:template match="text:span" mode="text-box">

    <xsl:variable name="tekst">
      <xsl:value-of select="."/>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$tekst != '' or text:s">
        <a:r>
          <a:rPr lang="en-EN" sz="1200">
            <xsl:for-each select="key('style',@text:style-name)/style:text-properties">
              <xsl:call-template name="InsertRunProperties"/>
            </xsl:for-each>
          </a:rPr>
          <a:t>
            <xsl:apply-templates mode="text-box"/>
          </a:t>
        </a:r>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertEndPara"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="text()[parent::text:p and parent::node()/text:span]" mode="text-box">
    <a:r>
      <a:rPr lang="en-EN" sz="1200">
        <xsl:for-each select="key('style',@text:style-name)/style:text-properties">
          <xsl:call-template name="InsertRunProperties"/>
        </xsl:for-each>

        <!-- default font face -->
        <xsl:if test="not(key('style',@text:style-name)/style:text-properties)">
          <a:latin typeface="Times New Roman" charset="0"/>
        </xsl:if>

      </a:rPr>
      <a:t>
        <xsl:value-of select="."/>
      </a:t>
    </a:r>
  </xsl:template>

  <xsl:template match="text()" mode="text-box">
    <xsl:value-of select="."/>
  </xsl:template>

  <xsl:template match="text:s" mode="text-box">

    <pxs:s xmlns:pxs="urn:cleverage:xmlns:post-processings:extra-spaces">
      <xsl:attribute name="pxs:c">
        <xsl:choose>
          <xsl:when test="@text:c">
            <xsl:value-of select="@text:c"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>1</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </pxs:s>

  </xsl:template>

  <xsl:template name="InsertEndPara">
    <a:endParaRPr lang="en-EN" sz="1400"/>
  </xsl:template>

  <xsl:template name="InsertRunProperties">
    <!-- used by text-boxes and charts -->

    <!-- font-size -->
    <xsl:if test="@fo:font-size">
      <xsl:attribute name="sz">
        <!-- in charts @fo:font-size can be a fraction so it is rounded -->
        <xsl:value-of select="number(round(substring-before(@fo:font-size,'pt')) * 100)"/>
      </xsl:attribute>
    </xsl:if>

    <!-- bold -->
    <xsl:if test="@fo:font-weight = 'bold' ">
      <xsl:attribute name="b">
        <xsl:text>1</xsl:text>
      </xsl:attribute>
    </xsl:if>

    <!-- italic -->
    <xsl:if test="@fo:font-style = 'italic' ">
      <xsl:attribute name="i">
        <xsl:text>1</xsl:text>
      </xsl:attribute>
    </xsl:if>

    <!-- strike-through -->
    <xsl:choose>
      <xsl:when test="@style:text-line-through-type = 'solid' ">
        <xsl:attribute name="strike">
          <xsl:text>sngStrike</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="contains(@style:text-line-through-type,'double' )">
        <xsl:attribute name="strike">
          <xsl:text>dblStrike</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- style:text-line-through-style-->
      <xsl:when test="@style:text-line-through-style = 'solid' ">
        <xsl:attribute name="strike">
          <xsl:text>sngStrike</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>

    <!-- font underline-->
    <xsl:choose>
      <!-- double-->
      <xsl:when
        test="@style:text-underline-style = 'solid' and contains(@style:text-underline-type, 'double' )">
        <xsl:attribute name="u">
          <xsl:text>dbl</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- bold -->
      <xsl:when
        test="@style:text-underline-style  = 'solid' and contains(@style:text-underline-width, 'bold' )">
        <xsl:attribute name="u">
          <xsl:text>heavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- single-->
      <xsl:when
        test="@style:text-underline-style = 'solid' and contains(@style:text-underline-width, 'auto' )">
        <xsl:attribute name="u">
          <xsl:text>sng</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- dotted -->
      <xsl:when
        test="@style:text-underline-style = 'dotted' and contains(@style:text-underline-width, 'auto' )">
        <xsl:attribute name="u">
          <xsl:text>dotted</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- bold dotted -->
      <xsl:when
        test="@style:text-underline-style = 'dotted' and contains(@style:text-underline-width, 'bold' )">
        <xsl:attribute name="u">
          <xsl:text>dottedHeavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- dash -->
      <xsl:when
        test="@style:text-underline-style = 'dash' and contains(@style:text-underline-width, 'auto' )">
        <xsl:attribute name="u">
          <xsl:text>dash</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- bold dash-->
      <xsl:when
        test="@style:text-underline-style = 'dash' and contains(@style:text-underline-width, 'bold' )">
        <xsl:attribute name="u">
          <xsl:text>dashHeavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- dash long -->
      <xsl:when
        test="@style:text-underline-style = 'long-dash' and contains(@style:text-underline-width, 'auto' )">
        <xsl:attribute name="u">
          <xsl:text>dashLong</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- bold dash long -->
      <xsl:when
        test="@style:text-underline-style = 'long-dash' and contains(@style:text-underline-width, 'bold' )">
        <xsl:attribute name="u">
          <xsl:text>dashLongHeavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- dot dash -->
      <xsl:when
        test="@style:text-underline-style = 'dot-dash' and contains(@style:text-underline-width, 'auto' )">
        <xsl:attribute name="u">
          <xsl:text>dotDash</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- bold dot dash -->
      <xsl:when
        test="@style:text-underline-style = 'dot-dash' and contains(@style:text-underline-width, 'bold' )">
        <xsl:attribute name="u">
          <xsl:text>dotDashHeavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- dot-dot-dash -->
      <xsl:when
        test="@style:text-underline-style= 'dot-dot-dash' and contains(@style:text-underline-width, 'auto' )">
        <xsl:attribute name="u">
          <xsl:text>dotDotDash</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- bold dot-dot-dash -->
      <xsl:when
        test="@style:text-underline-style= 'dot-dot-dash' and contains(@style:text-underline-width, 'bold' )">
        <xsl:attribute name="u">
          <xsl:text>dotDotDashHeavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- double wavy -->
      <xsl:when
        test="contains(@style:text-underline-style, 'wave' ) and contains(@style:text-underline-type, 'double' )">
        <xsl:attribute name="u">
          <xsl:text>wavyDbl</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- wavy -->
      <xsl:when
        test="contains(@style:text-underline-style, 'wave' ) and contains(@style:text-underline-width, 'auto' )">
        <xsl:attribute name="u">
          <xsl:text>wavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!-- bold wavy -->
      <xsl:when
        test="contains(@style:text-underline-style, 'wave' ) and contains(@style:text-underline-width, 'bold' )">
        <xsl:attribute name="u">
          <xsl:text>wavyHeavy</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>

    <!-- text position -->
    <xsl:if test="@style:text-position">
      <xsl:attribute name="baseline">
        <xsl:choose>
          <!-- superscript -->
          <xsl:when test="substring-before(@style:text-position,' ') = 'super' ">
            <xsl:value-of
              select="number(substring-before(substring-after(@style:text-position,' '), '%' )) * 1000"
            />
          </xsl:when>
          <!-- subscript -->
          <xsl:when test="substring-before(@style:text-position,' ') = 'sub' ">
            <xsl:text>-</xsl:text>
            <xsl:value-of
              select="number(substring-before(substring-after(@style:text-position,' '), '%' )) * 1000"
            />
          </xsl:when>
          <!-- when position is specified like 30% 75% -->
          <xsl:when test="contains(@style:text-position, ' ')">
            <xsl:value-of select="number(substring-before(@style:text-position,'%' )) * 1000"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- letter kerning-->
    <xsl:if test="style:text-properties/@style:letter-kerning = 'true' ">
      <xsl:attribute name="kern">
        <xsl:value-of select="100"/>
      </xsl:attribute>
    </xsl:if>

    <!--character spacing -->
    <xsl:if test="@fo:letter-spacing">
      <xsl:attribute name="spc">
        <!-- condensed -->
        <xsl:if test="substring-before(@fo:letter-spacing, 'cm' )&lt; 0 ">
          <xsl:value-of
            select="format-number(substring-before(@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
        </xsl:if>
        <!-- expanded -->
        <xsl:if test="substring-before(@fo:letter-spacing, 'cm' ) &gt;= 0">
          <xsl:value-of
            select="format-number((substring-before(@fo:letter-spacing, 'cm' ) * 72 div 2.54) *100 ,'#')"
          />
        </xsl:if>
      </xsl:attribute>
    </xsl:if>

    <!-- font color -->
    <xsl:if test="@fo:color">
      <a:solidFill>
        <a:srgbClr>
          <xsl:attribute name="val">
            <xsl:value-of select="substring-after(@fo:color,'#')"/>
          </xsl:attribute>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>

    <!-- text shadow -->
    <xsl:if test="@fo:text-shadow != 'none'">
      <a:effectLst>
        <a:outerShdw blurRad="38100" dist="38100" dir="2700000">
          <a:srgbClr val="000000">
            <a:alpha val="43137"/>
          </a:srgbClr>
        </a:outerShdw>
      </a:effectLst>
    </xsl:if>

    <xsl:if test="@style:text-underline-color">
      <a:uFill>
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(@style:text-underline-color,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </a:uFill>
    </xsl:if>

    <!-- font family  (Times New Roman default)-->
    <a:latin typeface="Times New Roman" charset="0">
      <xsl:if test="@fo:font-family">
        <xsl:attribute name="typeface">
          <xsl:value-of select="translate(@fo:font-family, &quot;'&quot;,'')"/>
        </xsl:attribute>
      </xsl:if>
    </a:latin>

  </xsl:template>

  <xsl:template name="InsertParagraphProperties">
    <xsl:for-each select="key('style',@text:style-name)/style:paragraph-properties">
      <a:pPr>

        <!-- first line indent -->
        <xsl:if test="@fo:text-indent and substring-before(@fo:text-indent,'cm') != 0">
          <xsl:attribute name="indent">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="@fo:text-indent"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>

        <!-- left margin -->
        <xsl:if test="@fo:margin-left and substring-before(@fo:margin-left,'cm') &gt; 0">
          <xsl:attribute name="marL">
            <!--fo:margin-left-->
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="@fo:margin-left"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>

        <!-- horizontal align -->
        <xsl:if test="@fo:text-align">
          <xsl:attribute name="algn">
            <xsl:choose>
              <xsl:when test="@fo:text-align = 'center' ">
                <xsl:text>ctr</xsl:text>
              </xsl:when>
              <xsl:when test="@fo:text-align = 'end' ">
                <xsl:text>r</xsl:text>
              </xsl:when>
              <xsl:when test="@fo:text-align = 'justify' ">
                <xsl:text>just</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>l</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>

        <!-- line spacing -->
        <!-- if the line spacing is in terms of Percentage, multiply the value with 1000 -->
        <xsl:if
          test="@fo:line-height and substring-before(@fo:line-height,'%') &gt; 0 and not(substring-before(@fo:line-height,'%') = 100)">
          <a:lnSpc>
            <a:spcPct>
              <xsl:attribute name="val">
                <xsl:value-of
                  select="format-number(substring-before(@fo:line-height,'%')* 1000,'#.##') "/>
              </xsl:attribute>
            </a:spcPct>
          </a:lnSpc>
        </xsl:if>
        <!--If the line spacing is in terms of Points,multiply the value with 2835-->
        <xsl:if test="@style:line-spacing and substring-before(@style:line-spacing,'cm') &gt; 0">
          <a:lnSpc>
            <a:spcPts>
              <xsl:attribute name="val">
                <xsl:value-of select="round(substring-before(@style:line-spacing,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:lnSpc>
        </xsl:if>
        <xsl:if
          test="@style:line-height-at-least and substring-before(@style:line-height-at-least,'cm') &gt; 0 ">
          <a:lnSpc>
            <a:spcPts>
              <xsl:attribute name="val">
                <xsl:value-of
                  select="round(substring-before(@style:line-height-at-least,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:lnSpc>
        </xsl:if>

        <!--spacing above -->
        <xsl:if test="@fo:margin-top and substring-before(@fo:margin-top,'cm') &gt; 0 ">
          <a:spcBef>
            <a:spcPts>
              <xsl:attribute name="val">
                <xsl:value-of select="round(substring-before(@fo:margin-top,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:spcBef>
        </xsl:if>

        <!-- spacing below -->
        <xsl:if test="@fo:margin-bottom and substring-before(@fo:margin-bottom,'cm') &gt; 0 ">
          <a:spcAft>
            <a:spcPts>
              <xsl:attribute name="val">
                <xsl:value-of select="round(substring-before(@fo:margin-bottom,'cm')* 2835) "/>
              </xsl:attribute>
            </a:spcPts>
          </a:spcAft>
        </xsl:if>

      </a:pPr>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertTextBoxProperties">
    <a:bodyPr wrap="square" rtlCol="0" anchor="t">
      <xsl:for-each select="key('style',@draw:style-name)/style:graphic-properties">

        <xsl:attribute name="lIns">
          <xsl:choose>
            <xsl:when test="@fo:padding-left">
              <xsl:call-template name="emu-measure">
                <xsl:with-param name="unit" select="'cm'"/>
                <xsl:with-param name="length" select="@fo:padding-left"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>0</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

        <xsl:attribute name="tIns">
          <xsl:choose>
            <xsl:when test="@fo:padding-top">
              <xsl:call-template name="emu-measure">
                <xsl:with-param name="unit" select="'cm'"/>
                <xsl:with-param name="length" select="@fo:padding-top"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>0</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

        <xsl:attribute name="rIns">
          <xsl:choose>
            <xsl:when test="@fo:padding-right">
              <xsl:call-template name="emu-measure">
                <xsl:with-param name="unit" select="'cm'"/>
                <xsl:with-param name="length" select="@fo:padding-right"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>0</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

        <xsl:attribute name="bIns">
          <xsl:choose>
            <xsl:when test="@fo:padding-bottom">
              <xsl:call-template name="emu-measure">
                <xsl:with-param name="unit" select="'cm'"/>
                <xsl:with-param name="length" select="@fo:padding-bottom"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>0</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

      </xsl:for-each>
    </a:bodyPr>

  </xsl:template>

  <xsl:template name="InsertDrawingFill">
    <xsl:param name="chartDirectory"/>

    <xsl:choose>
      <!-- background image -->
      <xsl:when test="@draw:fill = 'bitmap' ">
        <xsl:call-template name="InsertBitmapFill">
          <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
        </xsl:call-template>
      </xsl:when>
      <!-- background color-->
      <xsl:when test="not(@draw:fill = 'none' )">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(@draw:fill-color, '#')"/>
            </xsl:attribute>

            <!-- transparency -->
            <xsl:if test="@draw:opacity">
              <a:alpha>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-before(@draw:opacity,'%' ) * 1000"/>
                </xsl:attribute>
              </a:alpha>
            </xsl:if>

          </a:srgbClr>
        </a:solidFill>
      </xsl:when>
      <xsl:otherwise>
        <a:noFill/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertDrawingBorder">

    <!-- line style -->
    <a:ln>
      <xsl:attribute name="w">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="@svg:stroke-width"/>
          <xsl:with-param name="unit">emu</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>

      <xsl:choose>
        <xsl:when test="not(@draw:stroke = 'none' )">
          <a:solidFill>
            <a:srgbClr>

              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="@svg:stroke-color != '' ">
                    <xsl:value-of select="substring-after(@svg:stroke-color, '#')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>000000</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>

              <xsl:if test="@svg:stroke-opacity != ''">
                <xsl:variable name="transparency">
                  <xsl:value-of select="substring-before(@svg:stroke-opacity, '%')"/>
                </xsl:variable>
                <a:alpha>
                  <xsl:attribute name="val">
                    <xsl:value-of select="$transparency * 1000"/>
                  </xsl:attribute>
                </a:alpha>
              </xsl:if>

            </a:srgbClr>
          </a:solidFill>
        </xsl:when>
        <xsl:otherwise>
          <a:noFill/>
        </xsl:otherwise>
      </xsl:choose>

    </a:ln>
  </xsl:template>

  <xsl:template name="InsertBitmapFill">
    <xsl:param name="chartDirectory"/>

    <xsl:variable name="fillImage">
      <xsl:value-of select="@draw:fill-image-name"/>
    </xsl:variable>

    <xsl:for-each
      select="document(concat($chartDirectory,'/styles.xml'))/office:document-styles/office:styles/draw:fill-image[@draw:name = $fillImage]">
      <a:blipFill dpi="0" rotWithShape="1">
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          r:embed="{generate-id()}"/>
        <a:srcRect/>
        <a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>
      </a:blipFill>
    </xsl:for-each>

    <!--pzip:copy pzip:source="{@xlink:href}" pzip:target="xl/media/{$imageName}"/-->

    <!--xsl:for-each select="document()">
    
  </xsl:for-each-->
  </xsl:template>

</xsl:stylesheet>

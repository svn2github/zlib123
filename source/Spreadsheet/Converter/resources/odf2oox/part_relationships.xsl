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
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
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
        <xsl:if test="not(table:scenario)">
          <Relationship Id="{generate-id(.)}"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet">
            <xsl:variable name="NumberSheet">
              <xsl:value-of select="position()"/>
            </xsl:variable>
            <xsl:attribute name="Target">
              <xsl:value-of select="concat(concat('worksheets/sheet', $NumberSheet), '.xml')"/>
            </xsl:attribute>
          </Relationship>
        </xsl:if>
      </xsl:for-each>
      
      <xsl:if test="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:tracked-changes">
        <Relationship Id="{generate-id(.)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders" Target="revisions/revisionHeaders.xml" />
        <Relationship Id="{concat(generate-id(.), 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames" Target="revisions/userNames.xml" /> 
      </xsl:if>

      <!--  Static relationships -->
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        Target="styles.xml"/>
      <Relationship Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        Target="sharedStrings.xml"/>
      <Relationship Id="rId3"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
        Target="connections.xml"/>
      <xsl:call-template name="InsertWorkobookExternalRels"/>


    </Relationships>

  </xsl:template>
 
  
  <xsl:template name="InsertRevisionsRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:tracked-changes/table:cell-content-change">                
        <Relationship Id="{generate-id()}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog" Target="{concat('revisionLog', generate-id(), '.xml')}" /> 
      </xsl:for-each>
    </Relationships>
  </xsl:template>
  
  <xsl:template match="table:cell-content-change" mode="revisionRels">
    
    <Relationship Id="{generate-id()}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog" Target="{concat('revisionLog', generate-id(), '.xml')}" /> 
    
    <xsl:if test="following-sibling::node()[1][name()='table:cell-content-change']">
      <xsl:apply-templates select="following-sibling::node()[1][name()='table:cell-content-change']" mode="revisionRels"/>
    </xsl:if>
    
  </xsl:template>

  <xsl:template name="TranslateIllegalChars">
    <xsl:param name="string"/>

    <xsl:choose>

      <!-- remove space-->
      <xsl:when test="contains($string,' ')">
        <xsl:choose>
          <xsl:when test="substring-before($string,' ') =''">
            <xsl:call-template name="TranslateIllegalChars">
              <xsl:with-param name="string" select="substring-after($string,' ')"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="substring-before($string,' ') !=''">
            <xsl:call-template name="TranslateIllegalChars">
              <xsl:with-param name="string"
                select="concat(substring-before($string,' '),substring-after($string,' '))"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <!-- change  '&lt;' to '%3C'  after conversion-->
      <xsl:when test="contains($string,'&lt;')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'&lt;') =''">
            <xsl:call-template name="TranslateIllegalChars">
              <xsl:with-param name="string"
                select="concat('%3C',substring-after($string,'&lt;'))"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="substring-before($string,'&lt;') !=''">
            <xsl:call-template name="TranslateIllegalChars">
              <xsl:with-param name="string"
                select="concat(substring-before($string,'&lt;'),'%3C',substring-after($string,'&lt;'))"
              />
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <!-- change  '&gt;' to '%3E'  after conversion-->
      <xsl:when test="contains($string,'&gt;')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'&gt;') =''">
            <xsl:call-template name="TranslateIllegalChars">
              <xsl:with-param name="string"
                select="concat('%3E',substring-after($string,'&gt;'))"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="substring-before($string,'&gt;') !=''">
            <xsl:call-template name="TranslateIllegalChars">
              <xsl:with-param name="string"
                select="concat(substring-before($string,'&gt;'),'%3E',substring-after($string,'&gt;'))"
              />
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <xsl:otherwise>
        <xsl:value-of select="$string"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>


  <xsl:template name="InsertWorksheetsRels">
    <xsl:param name="sheetNum"/>
    <xsl:param name="comment"/>
    <xsl:param name="picture"/>
    <xsl:param name="hyperlink"/>
    <xsl:param name="chart"/>
    <xsl:param name="textBox"/>
    <xsl:param name="OLEObject"/>

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

      <!--hyperlink-->
      <xsl:if test="$hyperlink = 'true' ">
        <xsl:for-each select="descendant::text:a[not(ancestor::draw:custom-shape) and not(ancestor::office:annotation)]">

          <Relationship Id="{generate-id(.)}"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            TargetMode="External">

            <xsl:attribute name="Target">
              <xsl:if test="@xlink:href">
                <xsl:choose>
                  <!-- when hyperlink to a site or mailto-->
                  <xsl:when test="contains(@xlink:href,':') and not(starts-with(@xlink:href,'/'))">
                    <xsl:call-template name="TranslateIllegalChars">
                      <xsl:with-param name="string" select="@xlink:href"/>
                    </xsl:call-template>
                  </xsl:when>
                  <!-- when hyperlink to an document-->
                  <xsl:otherwise>

                    <xsl:variable name="translatedTarget">
                      <xsl:call-template name="SpaceTo20Percent">
                        <xsl:with-param name="string" select="@xlink:href"/>
                      </xsl:call-template>
                    </xsl:variable>

                    <xsl:choose>
                      <!-- when starts with up folder sign -->
                      <xsl:when test="starts-with($translatedTarget,'../' )">
                        <xsl:value-of
                          select="translate(substring-after($translatedTarget,'../'),'/','\')"/>
                      </xsl:when>
                      <!-- when file is in another disk -->
                      <xsl:when test="starts-with($translatedTarget,'/')">
                        <xsl:value-of
                          select="concat('file:///',translate(substring-after($translatedTarget,'/'),'/','\'))"
                        />
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="translate($translatedTarget,'/','\')"/>
                      </xsl:otherwise>
                    </xsl:choose>

                  </xsl:otherwise>
                </xsl:choose>
              </xsl:if>
            </xsl:attribute>
          </Relationship>
        </xsl:for-each>
      </xsl:if>

      <xsl:for-each
        select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell/table:cell-range-source">

        <Relationship Id="{generate-id()}"
          Target="{concat('../queryTables/queryTable', position(), '.xml')}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"/>

      </xsl:for-each>

      <!-- drawing.xml file -->
      <xsl:if test="contains($chart,'true') or $picture = 'true' or $textBox = 'true' ">
        <Relationship Id="{concat('d_rId',$sheetNum)}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
          Target="{concat('../drawings/drawing',$sheetNum,'.xml')}"/>
      </xsl:if>

      <xsl:if test="$OLEObject = 'true'">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
          Target="{concat('../drawings/vmlDrawing', $sheetNum,'.vml')}"/>
      </xsl:if>
    </Relationships>
  </xsl:template>


  <xsl:template name="InsertChartRels">
    <xsl:param name="chartDirectory"/>
    
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!-- search for bitmap fills -->
      <xsl:for-each
        select="/office:document-content/office:automatic-styles/style:style/style:graphic-properties[@draw:fill = 'bitmap' ]">
        
        <xsl:variable name="fillImage">
          <xsl:value-of select="@draw:fill-image-name"/>
        </xsl:variable>
        
        <!-- id based on a style in content.xml -->
        <xsl:variable name="fillId">
          <xsl:value-of select="generate-id()"/>
        </xsl:variable>
        
        <!-- go to bitmap-fill style -->
        <xsl:for-each
          select="document(concat($chartDirectory,'/styles.xml'))/office:document-styles/office:styles/draw:fill-image[@draw:name = $fillImage]">
          
          <xsl:variable name="imageName" select="substring-after(@xlink:href, 'Pictures/')"/>
          
          <pzip:copy pzip:source="{concat($chartDirectory,'/',@xlink:href)}" pzip:target="xl/media/{concat($fillId,$imageName)}"/>
          
          <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
            <xsl:attribute name="Id">
              <xsl:value-of select="$fillId"/>
            </xsl:attribute>
            
            <xsl:attribute name="Target">
              <xsl:text>../media/</xsl:text>
              <!-- the same bitmap fill in two diferent charts has the same name so to distinct them it is preceded with generate-id() -->
              <xsl:value-of select="concat($fillId,$imageName)"/>
            </xsl:attribute>
          </Relationship>
        </xsl:for-each>
        
      </xsl:for-each>
    </Relationships>    
  </xsl:template>
  
  <xsl:template name="InsertDrawingRels">
    <xsl:param name="sheetNum"/>

    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!-- chart rels -->
      <xsl:variable name="chart">
        <xsl:for-each select="descendant::draw:frame/draw:object">
          <xsl:choose>
            <xsl:when test="not(document(concat(translate(@xlink:href,'./',''),'/settings.xml')))">
              <xsl:for-each select="document(concat(translate(@xlink:href,'./',''),'/content.xml'))">
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

      <xsl:if test="contains($chart, 'true')">
        <xsl:for-each
          select="descendant::draw:frame/draw:object[document(concat(translate(@xlink:href,'./',''),'/content.xml'))/office:document-content/office:body/office:chart]">
          <Relationship Id="{generate-id(parent::node())}"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="{concat('../charts/chart',$sheetNum,'_',position(),'.xml')}"/>
        </xsl:for-each>
      </xsl:if>

      <!-- pictures -->
      <xsl:for-each
        select="descendant::draw:frame/draw:image[not(name(parent::node()/parent::node()) = 'draw:g' )  and not(parent::node()/draw:object)]">

        <xsl:choose>
          <!-- embeded pictures -->
          <xsl:when test="starts-with(@xlink:href, 'Pictures/')">
            <xsl:variable name="imageName" select="substring-after(@xlink:href, 'Pictures/')"/>
            <pzip:copy pzip:source="{@xlink:href}" pzip:target="xl/media/{$imageName}"/>
            <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
              Id="{generate-id(parent::node())}"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="../media/{$imageName}"/>
          </xsl:when>
          <!-- linked pictures -->
          <xsl:otherwise>

            <!-- change spaces to %20 -->
            <xsl:variable name="translatedTarget">
              <xsl:call-template name="TranslateIllegalChars">
                <xsl:with-param name="string" select="@xlink:href"/>
              </xsl:call-template>
            </xsl:variable>

            <xsl:variable name="target">
              <xsl:choose>
                <!-- when starts with up folder sign -->
                <xsl:when test="starts-with($translatedTarget,'../' )">
                  <xsl:value-of select="translate(substring-after($translatedTarget,'../'),'/','\')"
                  />
                </xsl:when>
                <!-- when file is in another disk -->
                <xsl:when test="starts-with($translatedTarget,'/')">
                  <xsl:value-of
                    select="concat('file:///',translate(substring-after($translatedTarget,'/'),'/','\'))"
                  />
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="translate($translatedTarget,'/','\')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>

            <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
              Id="{generate-id(parent::node())}"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="{$target}" TargetMode="External"/>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:for-each>

    </Relationships>
  </xsl:template>

  <xsl:template name="SpaceTo20Percent">
    <xsl:param name="string"/>

    <xsl:choose>
      <!-- change space to  '%20' after conversion-->
      <xsl:when test="contains($string,' ')">
        <xsl:choose>
          <xsl:when test="substring-before($string,' ') =''">
            <xsl:call-template name="TranslateIllegalChars">
              <xsl:with-param name="string" select="concat('%20',substring-after($string,' '))"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="substring-before($string,' ') !=''">
            <xsl:call-template name="SpaceTo20Percent">
              <xsl:with-param name="string"
                select="concat(substring-before($string,' '),'%20',substring-after($string,' '))"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <xsl:otherwise>
        <xsl:value-of select="$string"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertLinkExternalRels">

    <xsl:for-each
      select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell/table:cell-range-source">

      <xsl:call-template name="InsertQueryTable"/>

    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertWorkobookExternalRels">
    <xsl:if
      test="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:shapes/draw:frame/draw:object">
      <xsl:for-each
        select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:shapes/draw:frame">
        <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink">
          <xsl:attribute name="Id">
            <xsl:value-of select="generate-id()"/>
          </xsl:attribute>

          <xsl:variable name="NumberFile">
            <xsl:value-of select="position()"/>
          </xsl:variable>

          <xsl:attribute name="Target">
            <xsl:for-each select="draw:object">
              <xsl:value-of
                select="concat(concat('externalLinks/externalLink', $NumberFile), '.xml')"/>
            </xsl:for-each>
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  
  <!-- Insert Revision Headers -->
  <xsl:template name="revisionHeaderProperties">
    <headers xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      guid="{concat('{', '100', '}')}" diskRevisions="1" revisionId="2" version="3">          
      <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:tracked-changes">         
        <xsl:apply-templates select="node()[1][name()='table:cell-content-change']" mode="revisionHeaders">
          <xsl:with-param name="num">
            <xsl:text>1</xsl:text>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:for-each>
    </headers>
  </xsl:template>
  
  <xsl:template match="table:cell-content-change" mode="revisionHeaders">
    <xsl:param name="num"/>
    
    <header maxSheetId="4" r:id="{generate-id()}" guid="{concat('{', $num, '}')}">
      <xsl:attribute name="userName">
        <xsl:value-of select="office:change-info/dc:creator"/>
      </xsl:attribute>
      <xsl:attribute name="dateTime">
        <xsl:value-of select="office:change-info/dc:date"/>
      </xsl:attribute>
      <sheetIdMap>
      <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table">         
        <sheetId>
          <xsl:attribute name="val">
            <xsl:value-of select="position()"/>
          </xsl:attribute>
        </sheetId>
      </xsl:for-each>
      </sheetIdMap>
    </header>
    
    <xsl:if test="following-sibling::node()[1][name()='table:cell-content-change']">
      <xsl:apply-templates select="following-sibling::node()[1][name()='table:cell-content-change']" mode="revisionHeaders">
        <xsl:with-param name="num">
          <xsl:value-of select="$num + 1"/>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:if>
    
  </xsl:template>
  
  

</xsl:stylesheet>

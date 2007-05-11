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
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">
  
  <xsl:key name="drawing" match="e:drawing" use="''"/>
  
  <!-- We check cell when the picture is starting and ending -->
  <xsl:template name="PictureCell">
    <xsl:param name="sheet"/>    
    <xsl:apply-templates select="e:worksheet/e:drawing">
     <xsl:with-param name="sheet">
       <xsl:value-of select="$sheet"/>
     </xsl:with-param>
    </xsl:apply-templates>
  </xsl:template>
  
  <!-- Get Row with Picture -->
  <xsl:template name="PictureRow">
    <xsl:param name="PictureCell"/>
    <xsl:param name="Result"/>    
    <xsl:choose>
      <xsl:when test="$PictureCell != ''">        
        <xsl:call-template name="PictureRow">          
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="substring-after($PictureCell, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="Result">
            <xsl:value-of select="concat($Result,  concat(substring-before($PictureCell, ':'), ';'))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$Result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="e:drawing">
    <xsl:param name="sheet"/>
    
  <xsl:call-template name="GetTargetPicture">
    <xsl:with-param name="sheet">
      <xsl:value-of select="$sheet"/>
    </xsl:with-param>
    <xsl:with-param name="id">
      <xsl:value-of select="@r:id"/>
    </xsl:with-param>
  </xsl:call-template>
    
  </xsl:template>
  
  <!-- We check drawing's file -->
  <xsl:template name="GetTargetPicture">    
    <xsl:param name="id"/>
    <xsl:param name="sheet"/>
    <xsl:if test="document(concat(concat('xl/worksheets/_rels/', $sheet), '.rels'))//node()[name()='Relationship']">
      <xsl:for-each
        select="document(concat(concat('xl/worksheets/_rels/', $sheet), '.rels'))//node()[name()='Relationship']">
        <xsl:if test="./@Id=$id">
          <xsl:for-each select="document(concat('xl/', substring-after(@Target, '/')))">
            <xsl:if test="xdr:wsDr/xdr:twoCellAnchor">
              <xsl:apply-templates select="xdr:wsDr/xdr:twoCellAnchor[1]"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>  
  
  
  <!-- We check cell when the picture is starting and ending -->
  
  <xsl:template match="xdr:twoCellAnchor">
    <xsl:param name="PictureCell"/>
    
    <xsl:variable name="PictureColStart">
      <xsl:value-of select="xdr:from/xdr:col"/>
    </xsl:variable>
    
    <xsl:variable name="PictureRowStart">
      <xsl:value-of select="xdr:from/xdr:row"/>
    </xsl:variable>    
        
    <xsl:choose>
      <xsl:when test="following-sibling::xdr:twoCellAnchor">
        <xsl:apply-templates select="following-sibling::xdr:twoCellAnchor[1]">
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="concat(concat(concat(concat($PictureCell, $PictureRowStart), ':'), $PictureColStart), ';')"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat(concat(concat(concat($PictureCell, $PictureRowStart), ':'), $PictureColStart), ';')"/>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>
  
  <!-- Insert Empty Rows before picture -->
  <xsl:template name="InsertEmptyRows">
    <xsl:param name="repeat"/>
    <xsl:param name="sheet"/>
    <xsl:if test="$repeat &gt; 0">
    <xsl:for-each select="document(concat('xl/',$sheet))">
      <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
        table:number-rows-repeated="{$repeat}">
        <table:table-cell table:number-columns-repeated="256"/>
      </table:table-row>
    </xsl:for-each>
    </xsl:if>
  </xsl:template>
  
  <!-- Insert Empty Cols before picture -->
  <xsl:template name="InsertEmptyColls">
    <xsl:param name="repeat"/>    
    <table:table-cell table:number-columns-repeated="{$repeat}"/>
  </xsl:template>
  
  <!-- Insert picture -->
  
  <xsl:template name="InsertPicture">
    <xsl:param name="rowNum"/>
    <xsl:param name="collNum"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="CollsWithPicture"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="Drawing"/>

  <xsl:if test="$CollsWithPicture != ''">
    
    <xsl:variable name="CollStart">
      <xsl:value-of select="number(xdr:from/xdr:col) - 1"/>
    </xsl:variable>   
    
    
    <xsl:variable name="id">
      <xsl:for-each select="document(substring-after($sheet, '/'))">
      <xsl:value-of select="key('drawing', '')/@r:id"/>
      </xsl:for-each>
    </xsl:variable>

    
    <xsl:if test="$collNum != number(substring-before($CollsWithPicture, ';'))">
    <xsl:call-template name="InsertEmptyColls">
      <xsl:with-param name="repeat">
        <xsl:value-of select="number(substring-before($CollsWithPicture, ';')) - $collNum"/>
      </xsl:with-param>
    </xsl:call-template>
    </xsl:if>

    <table:table-cell>
      
      <xsl:for-each
        select="document(concat(concat('xl/worksheets/_rels/', substring-after($sheet, '/')), '.rels'))//node()[name()='Relationship']">        
          <xsl:call-template name="CopyPictures">            
            <xsl:with-param name="document">
              <xsl:value-of select="concat($Drawing, '.rels')"/>
            </xsl:with-param>
            <xsl:with-param name="targetName">
              <xsl:text>Pictures</xsl:text>
            </xsl:with-param>
          </xsl:call-template>
       </xsl:for-each>
      
      <xsl:for-each select="xdr:wsDr/xdr:twoCellAnchor">
        <xsl:if test="xdr:from/xdr:col = number(substring-before($CollsWithPicture, ';')) and xdr:from/xdr:row = $rowNum">
          <draw:frame draw:z-index="0" draw:name="Graphics 1" draw:style-name="gr1"
            draw:text-style-name="P1" svg:width="4.907cm" svg:height="4.622cm" >
            <!--size-->
            <!--xsl:call-template name="SetSize"/-->
            
            <xsl:call-template name="SetPosition">
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
            </xsl:call-template>
            
            <draw:image xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
              <xsl:call-template name="InsertImageHref">
                <xsl:with-param name="document">
                  <xsl:value-of select="concat($Drawing, '.rels')"/>
                </xsl:with-param>
              </xsl:call-template>
              
              <text:p/>
            </draw:image>
          </draw:frame>
          
          
        </xsl:if>
      </xsl:for-each>
    <!--draw:frame table:end-cell-address="Sheet1.N65" table:end-x="0.287cm"
      table:end-y="0.281cm" draw:z-index="1" draw:name="Graphics 2" draw:style-name="gr1"
      draw:text-style-name="P1" svg:width="2.166cm" svg:height="3.874cm" svg:x="0"
      svg:y="0">
      <draw:image xlink:href="Pictures/100000000000032000000258B0234CE5.jpg"
        xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
        <text:p/>
      </draw:image>
    </draw:frame-->
      
      
    </table:table-cell>    
 
    <!-- Insert Next Picture in this row (if exist) -->
    
    <xsl:if test="substring-after($CollsWithPicture, ';') != '' and $CollsWithPicture != ''">
      <xsl:call-template name="InsertPicture">
        <xsl:with-param name="rowNum">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="collNum">
          <xsl:value-of select="substring-before($CollsWithPicture, ';') + 1"/>
        </xsl:with-param>
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="$PictureCell"/>
        </xsl:with-param>
        <xsl:with-param name="CollsWithPicture">
          <xsl:value-of select="substring-after($CollsWithPicture, ';')"/>
        </xsl:with-param>
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
        <xsl:with-param name="Drawing">
          <xsl:value-of select="$Drawing"/>
        </xsl:with-param>
        <xsl:with-param name="NameSheet">
          <xsl:value-of select="$NameSheet"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
  </xsl:if>  
  </xsl:template>
 
  
  <!-- Get min. row number with picture -->
  <xsl:template name="GetMinRowWithPicture">
    <xsl:param name="min"/>
    <xsl:param name="PictureRow"/>
    
    <xsl:variable name="numRow">
      <xsl:value-of select="substring-before($PictureRow, ';')"/>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="$PictureRow = ''">
        <xsl:value-of select="$min"/>
      </xsl:when>
      <xsl:when test="$min = ''">
        <xsl:call-template name="GetMinRowWithPicture">
          <xsl:with-param name="min">
            <xsl:value-of select="substring-before($PictureRow, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="substring-after($PictureRow, ';')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$min &gt; substring-before($PictureRow, ';')">
        <xsl:call-template name="GetMinRowWithPicture">
          <xsl:with-param name="min">
            <xsl:value-of select="substring-before($PictureRow, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="substring-after($PictureRow, ';')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="GetMinRowWithPicture">
          <xsl:with-param name="min">
            <xsl:value-of select="$min"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="substring-after($PictureRow, ';')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>
  
  <!-- Get colls with picture from this row  -->
  
  <xsl:template name="GetCollsWithPicture">
    <xsl:param name="rowNumber"/>
    <xsl:param name="PictureColl"/>
    <xsl:param name="PictureCell"/>

    <xsl:choose>
      <xsl:when test="contains($PictureCell, concat(concat(';',$rowNumber),':'))">
        <xsl:call-template name="GetCollsWithPicture">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$rowNumber"/>
          </xsl:with-param>
          <xsl:with-param name="PictureColl">
            <xsl:value-of select="concat($PictureColl, concat(substring-before(substring-after($PictureCell, concat(';', concat($rowNumber, ':'))),';'), ';'))"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="concat(';', substring-after(substring-after($PictureCell, concat(';', concat($rowNumber, ':'))),';'))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <!--xsl:when test="contains($PictureCell, concat(concat(';',$rowNumber),':'))">
        <xsl:call-template name="GetCollsWithPicture">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$rowNumber"/>
          </xsl:with-param>
          <xsl:with-param name="PictureColl">
            <xsl:value-of select="concat($PictureColl, concat(substring-before(substring-after($PictureCell, concat(';', concat($rowNumber, ':'))),';'), ';'))"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="concat($PictureCell, concat(';', substring-after(substring-after($PictureCell, concat(';', concat($rowNumber, ':'))),';')))"/>
          </xsl:with-param>
          </xsl:call-template>        
      </xsl:when-->
      <xsl:otherwise>
        <xsl:value-of select="$PictureColl"/>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>
 
  <!-- Insert Table Body -->
  
  <xsl:template name="TwoCellAnchor">
    <xsl:param name="prevRow" select="0"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="Drawing"/>

    <xsl:if test="$PictureRow != ''">
   
    <xsl:variable name="TableStyleName">
      <xsl:for-each select="document(concat('xl/',$sheet))">
        <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
      </xsl:for-each>
    </xsl:variable>
      
      <xsl:variable name="GetMinRowWithPicture">
        <xsl:call-template name="GetMinRowWithPicture">
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>      
     
    <xsl:call-template name="InsertEmptyRows">
       <xsl:with-param name="repeat">
          <xsl:value-of select="number($GetMinRowWithPicture) - number($prevRow)"/>
      </xsl:with-param>
      <xsl:with-param name="sheet">
         <xsl:value-of select="$sheet"/>
      </xsl:with-param>
    </xsl:call-template>
      
      <xsl:variable name="CollsWithPicture">
        <xsl:call-template name="GetCollsWithPicture">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$GetMinRowWithPicture"/>        
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="concat(';', $PictureCell)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>      

      <!-- Insert Empty Row with picture (pictures) -->
    <table:table-row table:style-name="{$TableStyleName}">
      <xsl:call-template name="InsertPicture">
       <xsl:with-param name="collNum">
          <xsl:text>0</xsl:text>
       </xsl:with-param>
        <xsl:with-param name="rowNum">
          <xsl:value-of select="$GetMinRowWithPicture"/>
        </xsl:with-param>
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="$PictureCell"/>
        </xsl:with-param>       
        <xsl:with-param name="CollsWithPicture">
          <xsl:value-of select="$CollsWithPicture"/>
        </xsl:with-param>
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
        <xsl:with-param name="Drawing">
          <xsl:value-of select="$Drawing"/>
        </xsl:with-param>   
        <xsl:with-param name="NameSheet">
          <xsl:value-of select="$NameSheet"/>
        </xsl:with-param>
      </xsl:call-template>
    </table:table-row>      
    
      
      <xsl:if test="concat(substring-after($PictureRow, concat($GetMinRowWithPicture,';')), substring-before($PictureRow, concat($GetMinRowWithPicture,';'))) != ''">
        <xsl:call-template name="TwoCellAnchor">
          <xsl:with-param name="prevRow">
            <xsl:value-of select="$GetMinRowWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:call-template name="DeleteRow">
              <xsl:with-param name="GetMinRowWithPicture">
                <xsl:value-of select="$GetMinRowWithPicture"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="concat(substring-after($PictureRow, concat($GetMinRowWithPicture,';')), substring-before($PictureRow, concat($GetMinRowWithPicture,';')))"/>        
              </xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="Drawing">
            <xsl:value-of select="$Drawing"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:if>
    
    
  </xsl:template> 
  
  <!-- delete cell with picture which are inserted -->
  
  <xsl:template name="DeleteRow">
    <xsl:param name="PictureRow"/>
    <xsl:param name="GetMinRowWithPicture"/>
    <xsl:choose>
      <xsl:when test="contains($PictureRow, concat(';', concat($GetMinRowWithPicture,';'))) and $PictureRow != ''">
        <xsl:call-template name="DeleteRow">
          <xsl:with-param name="GetMinRowWithPicture">
            <xsl:value-of select="$GetMinRowWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="concat(substring-after($PictureRow, concat($GetMinRowWithPicture,';')), substring-before($PictureRow, concat($GetMinRowWithPicture,';')))"/>        
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$PictureRow"/>
      </xsl:otherwise>
      
    
    </xsl:choose>
    
    
      
    
  </xsl:template>
  
  
  <!-- Insert Empty Sheet with picture -->
  
  <xsl:template name="InsertEmptySheetWithPicture">
    <xsl:param name="PictureCell"/>    
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>

    <xsl:variable name="PictureRow">
    <xsl:call-template name="PictureRow">
     <xsl:with-param name="PictureCell">
        <xsl:value-of select="$PictureCell"/>
      </xsl:with-param>
    </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="id">
      <xsl:value-of select="key('drawing', '')/@r:id"/>        
    </xsl:variable>
    <xsl:for-each
      select="document(concat(concat('xl/worksheets/_rels/', substring-after($sheet, '/')), '.rels'))//node()[name()='Relationship']">
      <xsl:variable name="Drawing">
        <xsl:value-of select="substring-after(substring-after(@Target, '/'), '/')"/>
      </xsl:variable>
      <xsl:if test="./@Id=$id">
        <xsl:for-each select="document(concat('xl/', substring-after(@Target, '/')))">
          <xsl:if test="xdr:wsDr/xdr:twoCellAnchor">
            <xsl:call-template name="TwoCellAnchor">
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="$PictureRow"/>
              </xsl:with-param>
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="$PictureCell"/>
              </xsl:with-param>
              <xsl:with-param name="Drawing">
                <xsl:value-of select="$Drawing"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
            </xsl:call-template>           
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:for-each>
    
  </xsl:template>
  
 
  <!-- inserts image href from relationships -->
  <xsl:template name="InsertImageHref">
    <xsl:param name="document"/>
    <xsl:param name="rId"/>
    <xsl:param name="targetName"/>
    <xsl:param name="srcFolder" select="'Pictures'"/>
    
    <xsl:variable name="id">
      <xsl:choose>
        <xsl:when test="$rId != ''">
          <xsl:value-of select="$rId"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="xdr:pic/xdr:blipFill/a:blip/@r:embed"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:for-each
      select="document(concat('xl/drawings/_rels/',$document))//node()[name() = 'Relationship']">
      <xsl:if test="./@Id=$id">
      <xsl:variable name="targetmode">
      <xsl:value-of select="./@TargetMode"/>
      </xsl:variable>
      <xsl:variable name="pzipsource">
      <xsl:value-of select="./@Target"/>
      </xsl:variable>
      <xsl:variable name="pziptarget">
      <xsl:choose>
      <xsl:when test="$targetName != ''">
      <xsl:value-of select="$targetName"/>
      </xsl:when>
      <xsl:otherwise>
      <xsl:value-of select="substring-after(substring-after($pzipsource,'/'), '/')"/>
      </xsl:otherwise>
      </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="xlink:href">
      <xsl:choose>
      <xsl:when test="$targetmode='External'">
      <xsl:value-of select="$pziptarget"/>
      </xsl:when>
      <xsl:otherwise>
      <xsl:value-of select="concat($srcFolder,'/', $pziptarget)"/>
      </xsl:otherwise>
      </xsl:choose>
      </xsl:attribute>
      </xsl:if>
      </xsl:for-each>    
  </xsl:template>
  
  <xsl:template name="SetSize">
    <!--xsl:choose>
      <xsl:when test="a:graphic/a:graphicData/pic:pic/pic:spPr/a:ln">
        <xsl:variable name="border">
          <xsl:call-template name="ConvertEmu3">
            <xsl:with-param name="length">
              <xsl:value-of select="a:graphic/a:graphicData/pic:pic/pic:spPr/a:ln/@w"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="height">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="wp:extent/@cy"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="width">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="wp:extent/@cx"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="svg:height">
          <xsl:value-of
            select="concat(substring-before($height,'cm')+substring-before($border,'cm')+substring-before($border,'cm'),'cm')"
          />
        </xsl:attribute>
        <xsl:attribute name="svg:width">
          <xsl:value-of
            select="concat(substring-before($width,'cm')+substring-before($border,'cm')+substring-before($border,'cm'),'cm')"
          />
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise-->
        <!--xsl:attribute name="svg:height">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="xdr:from/xdr:colOff"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
          </xsl:attribute-->
   
        <!--xsl:attribute name="svg:width">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="xdr:from/xdr:rowOff"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
        </xsl:attribute-->
      <!--/xsl:otherwise>
    </xsl:choose-->
  </xsl:template>
  
  <xsl:template name="SetPosition">
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:attribute name="table:end-cell-address">
      <xsl:variable name="ColEnd">
        <xsl:call-template name="NumbersToChars">
          <xsl:with-param name="num">
            <xsl:value-of select="xdr:to/xdr:col"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="RowEnd">
        <xsl:value-of select="xdr:to/xdr:row"/>
      </xsl:variable>
      <xsl:value-of select="concat($NameSheet, '.', $ColEnd, $RowEnd)"/>
    </xsl:attribute>
    <xsl:attribute name="svg:x">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:from/xdr:colOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:y">     
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="xdr:from/xdr:rowOff"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>        
    </xsl:attribute>
    <xsl:attribute name="table:end-x">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:to/xdr:colOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="table:end-y">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:to/xdr:rowOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>        
    </xsl:attribute>
  </xsl:template>
  
  
</xsl:stylesheet>

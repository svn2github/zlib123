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
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  exclude-result-prefixes="style text office">

  <xsl:preserve-space elements="number:text"/>

  <xsl:key name="date-style" match="number:date-style" use="@style:name"/>

  <xsl:template match="text:page-number" mode="paragraph">
    <xsl:if test="@text:page-adjust">
      <xsl:message terminate="no">feedback:Page number offset</xsl:message>
    </xsl:if>
    <w:fldSimple w:instr=" PAGE ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <!--  STATISTICS FIELDS  -->
  <xsl:template match="text:page-count" mode="paragraph">
    <w:fldSimple w:instr=" NUMPAGES ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>
  
  <xsl:template match="text:word-count|text:character-count|text:paragraph-count " mode="paragraph">
    <w:fldSimple>
      <xsl:attribute name="w:instr">
        <xsl:choose>
          <xsl:when test="../text:word-count">
            NUMWORDS
          </xsl:when>
          <xsl:when test="../text:character-count">
            NUMCHARS
          </xsl:when>
          <xsl:when test="../text:paragraph-count ">
            DOCPROPERTY  Paragraphs
          </xsl:when>
        </xsl:choose>
      <xsl:call-template name="GetNumberFormattingSwitch"/>
        \* MERGEFORMAT
      </xsl:attribute>
      <xsl:apply-templates mode="paragraph"></xsl:apply-templates>      
    </w:fldSimple>
  </xsl:template>

  <!-- Date Fields -->
  <xsl:template match="text:date|text:creation-date|text:print-date|text:modification-date"
    mode="paragraph">
    <xsl:choose>
      <xsl:when test="@text:fixed='true'">
        <w:r>
          <w:t>
            <xsl:value-of select="text()"/>
          </w:t>
        </w:r>
      </xsl:when>
      <xsl:otherwise>
        <w:fldSimple>
          <xsl:if test="@number:automatic-order='true' ">
            <xsl:message terminate="no">feedback:Date format</xsl:message>
          </xsl:if>
          <xsl:variable name="curStyle" select="@style:data-style-name"/>
          <xsl:variable name="fieldType">
            <xsl:choose>
              <xsl:when test="self::text:creation-date">CREATEDATE</xsl:when>
              <xsl:when test="self::text:print-date">PRINTDATE</xsl:when>
              <xsl:when test="self::text:modification-date">SAVEDATE</xsl:when>
              <xsl:otherwise>DATE</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="dataStyle">
            <xsl:apply-templates
              select="/*/office:automatic-styles/number:date-style[@style:name=$curStyle]"
              mode="dataStyle"/>
          </xsl:variable>
          <xsl:attribute name="w:instr">
            <xsl:value-of select="concat($fieldType,' \@ &quot;',$dataStyle,'&quot;')"/>
          </xsl:attribute>
          <w:r>
            <w:rPr>
              <w:noProof/>
            </w:rPr>
            <w:t>
              <xsl:value-of select="text()"/>
            </w:t>
          </w:r>
        </w:fldSimple>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:text" mode="dataStyle">
    <xsl:value-of select="."/>
  </xsl:template>

  <xsl:template match="number:day-of-week" mode="dataStyle">
    <xsl:choose>
      <xsl:when test="@number:style='long'">dddd</xsl:when>
      <xsl:otherwise>ddd</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:day" mode="dataStyle">
    <xsl:choose>
      <xsl:when test="@number:style='long'">dd</xsl:when>
      <xsl:otherwise>d</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:month" mode="dataStyle">
    <xsl:if test="@number:textual='true'">MM</xsl:if>
    <xsl:choose>
      <xsl:when test="@number:style='long'">MM</xsl:when>
      <xsl:otherwise>M</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:year" mode="dataStyle">
    <xsl:choose>
      <xsl:when test="@number:style='long'">yyyy</xsl:when>
      <xsl:otherwise>yy</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Time Fields -->
  <xsl:template
    match="text:time|text:creation-time|text:editing-duration|text:print-time|text:modification-time"
    mode="paragraph">
    <xsl:choose>
      <xsl:when test="@text:fixed='true'">
        <w:r>
          <w:t>
            <xsl:value-of select="text()"/>
          </w:t>
        </w:r>
      </xsl:when>
      <xsl:otherwise>
        <w:fldSimple>
          <xsl:variable name="curStyle" select="@style:data-style-name"/>
          <xsl:variable name="fieldType">
            <xsl:choose>
              <xsl:when test="self::text:creation-time">CREATEDATE</xsl:when>
              <xsl:when test="self::text:editing-duration">EDITTIME</xsl:when>
              <xsl:when test="self::text:print-time">PRINTDATE</xsl:when>
              <xsl:when test="self::text:modification-time">SAVEDATE</xsl:when>
              <xsl:otherwise>TIME</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="dataStyle">
            <xsl:apply-templates
              select="/*/office:automatic-styles/number:time-style[@style:name=$curStyle]"
              mode="dataStyle"/>
          </xsl:variable>
          <xsl:attribute name="w:instr">
            <xsl:value-of select="concat($fieldType,' \@ &quot;',$dataStyle,'&quot;')"/>
          </xsl:attribute>
          <w:r>
            <w:rPr>
              <w:noProof/>
            </w:rPr>
            <w:t>
              <xsl:value-of select="text()"/>
            </w:t>
          </w:r>
        </w:fldSimple>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:hours" mode="dataStyle">
    <xsl:choose>
      <xsl:when test="parent::node()/number:am-pm">
        <xsl:choose>
          <xsl:when test="@number:style='long'">hh</xsl:when>
          <xsl:otherwise>h</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="@number:style='long'">HH</xsl:when>
          <xsl:otherwise>H</xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:minutes" mode="dataStyle">
    <xsl:choose>
      <xsl:when test="@number:style='long'">mm</xsl:when>
      <xsl:otherwise>m</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:seconds" mode="dataStyle">
    <xsl:choose>
      <xsl:when test="@number:style='long'">ss</xsl:when>
      <xsl:otherwise>s</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="number:am-pm" mode="dataStyle">am/pm</xsl:template>

  <!-- Author Fields -->
  <!-- TODO : comment csv file -->
  <xsl:template match="text:author-name[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" USERNAME \* MERGEFORMAT ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:author-initials[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" USERINITIALS \* Upper  \* MERGEFORMAT ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>
  
  <!-- User Fields -->
  <!-- TODO : comment csv file -->
  <xsl:template match="text:initial-creator[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" AUTHOR ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:creator[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" LASTSAVEDBY ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:description[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" COMMENTS ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>
  
  <xsl:template match="text:subject[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" SUBJECT ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:keywords[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" KEYWORDS ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:title[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" TITLE ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <!-- Sender Fields -->
  <xsl:template match="text:sender-firstname[not(@text:fixed='true')]|text:sender-lastname[not(@text:fixed='true')]" mode="paragraph">    
      <xsl:variable name="username">
        <xsl:value-of select="."/>
      </xsl:variable>
    <w:fldSimple>
      <xsl:attribute name="w:instr">
        <xsl:value-of select="concat('USERNAME ' ,$username,'\* MERGEFORMAT')"/>
      </xsl:attribute>
      <xsl:apply-templates mode="paragraph"></xsl:apply-templates>    
    </w:fldSimple>
  </xsl:template>
  
  <xsl:template match="text:sender-initials[not(@text:fixed='true')]" mode="paragraph">
    <xsl:variable name="userinitial">
      <xsl:value-of select="."/>
    </xsl:variable>
    <w:fldSimple>
      <xsl:attribute name="w:instr">
        <xsl:value-of select="concat('USERINITIALS ' ,$userinitial,'\* MERGEFORMAT')"/>
      </xsl:attribute>
      <xsl:apply-templates mode="paragraph"></xsl:apply-templates>    
    </w:fldSimple>
  </xsl:template>
  
  <xsl:template match="text:sender-street[not(@text:fixed='true')]|text:sender-country[not(@text:fixed='true')]|text:sender-postal-code[not(@text:fixed='true')]|text:sender-city[not(@text:fixed='true')]" mode="paragraph">
    <xsl:variable name="adress">
      <xsl:value-of select="."/>
    </xsl:variable>
    <w:fldSimple>
      <xsl:attribute name="w:instr">
        <xsl:value-of select="concat('USERADDRESS ' ,$adress,'\* MERGEFORMAT')"/>
      </xsl:attribute>
      <xsl:apply-templates mode="paragraph"></xsl:apply-templates>      
    </w:fldSimple>
  </xsl:template>  
  
  <xsl:template match="text:sender-title" mode="paragraph">
    <xsl:variable name="title">
      <xsl:value-of select="."/>
    </xsl:variable>
    <w:fldSimple>
      <xsl:attribute name="w:instr">
        <xsl:value-of select="concat('TITLE ' ,$title,'\* MERGEFORMAT')"/>
      </xsl:attribute>     
    </w:fldSimple>
  </xsl:template> 
  
  <xsl:template match="text:sender-company" mode="paragraph">
    <w:fldSimple w:instr=" DOCPROPERTY  Company  \* MERGEFORMAT ">
      <xsl:apply-templates mode="paragraph"></xsl:apply-templates>      
    </w:fldSimple>
  </xsl:template>
  
  <!-- Templates Fields -->
  <xsl:template match="text:template-name" mode="paragraph">
  <w:fldSimple w:instr=" TEMPLATE   \* MERGEFORMAT ">
    <xsl:apply-templates mode="paragraph"></xsl:apply-templates>
   </w:fldSimple>
  </xsl:template>
  
  
  <xsl:template match="text:editing-cycles[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" REVNUM ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <!-- chapter field. -->
  <xsl:template match="text:chapter" mode="paragraph">
    <!-- if field displays name, convert into a reference to default heading style. -->
    <xsl:if test="@text:display='name' and @text:outline-level">
      <xsl:variable name="outline-level" select="@text:outline-level"/>
      <!-- COMMENT : if the style changes name in the application, it may not be found and cause an error. -->
      <xsl:variable name="style">
        <xsl:value-of
          select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:default-outline-level=$outline-level]/@style:display-name"
        />
      </xsl:variable>
      <w:fldSimple>
        <xsl:attribute name="w:instr">
          <xsl:value-of select="concat('STYLEREF &quot;',$style,'&quot; \*MERGEFORMAT')"/>
        </xsl:attribute>
        <w:r>
          <w:rPr>
            <w:noProof/>
          </w:rPr>
          <xsl:apply-templates mode="text"/>
        </w:r>
      </w:fldSimple>
    </xsl:if>
  </xsl:template>

  <!--numbering type for sequence-->
  <xsl:template name="InsertSequenceFieldNumType">
    <xsl:variable name="numType">
      <xsl:call-template name="GetNumberFormattingSwitch"/>
    </xsl:variable>
    <xsl:attribute name="w:instr">
      <xsl:value-of select="concat('SEQ ', @text:name,' ', $numType)"/>
    </xsl:attribute>
  </xsl:template>
  
  <!-- file name fields-->
  <xsl:template match="text:file-name" mode="paragraph">
    <w:fldSimple w:instr="FILENAME   \* MERGEFORMAT">      
      <xsl:apply-templates mode="paragraph"></xsl:apply-templates>
    </w:fldSimple>
  </xsl:template>
  
  <!-- report lost fields -->
  <xsl:template match="text:description" mode="paragraph">
    <xsl:message terminate="no">feedback:description field</xsl:message>
  </xsl:template>

  <xsl:template match="text:printed-by" mode="paragraph">
    <xsl:message terminate="no">feedback:Printed-by field</xsl:message>
  </xsl:template>
  
</xsl:stylesheet>

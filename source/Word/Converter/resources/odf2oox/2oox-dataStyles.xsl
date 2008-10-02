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
  exclude-result-prefixes="style text office number">

  <xsl:key name="date-style" match="number:date-style" use="@style:name"/>


  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->
  
  <xsl:template match="text:page-number" mode="paragraph">
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <xsl:choose>
      <xsl:when test="@text:page-adjust">
        <w:r>
          <w:instrText xml:space="preserve"> =</w:instrText>
        </w:r>
        <w:fldSimple w:instr=" PAGE "/>
        <w:r>
          <w:instrText xml:space="preserve"> +<xsl:value-of select="@text:page-adjust"/> </w:instrText>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="separate"/>
        </w:r>
      </xsl:when>
      <xsl:otherwise>
        <w:r>
          <xsl:call-template name="InsertRunProperties"/>
          <w:instrText xml:space="preserve">PAGE </w:instrText>
          <xsl:if test="@style:num-format">
            <w:instrText>
              <xsl:call-template name="GetNumberFormattingSwitch"/>
            </w:instrText>
          </xsl:if>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="separate"/>
        </w:r>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:apply-templates mode="paragraph"/>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </xsl:template>

  <xsl:template match="text:page-count" mode="paragraph">
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <xsl:call-template name="InsertRunProperties"/>
      <w:instrText xml:space="preserve">NUMPAGES </w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <xsl:apply-templates mode="text"/>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </xsl:template>

  <xsl:template match="text:word-count|text:character-count|text:paragraph-count " mode="paragraph">
    <w:fldSimple>
      <xsl:attribute name="w:instr">
        <xsl:choose>
          <xsl:when test="../text:word-count"> NUMWORDS </xsl:when>
          <xsl:when test="../text:character-count"> NUMCHARS </xsl:when>
          <xsl:when test="../text:paragraph-count "> DOCPROPERTY Paragraphs </xsl:when>
        </xsl:choose>
        <xsl:call-template name="GetNumberFormattingSwitch"/> \* MERGEFORMAT </xsl:attribute>
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:date|text:creation-date|text:print-date|text:modification-date" mode="paragraph">
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
            <xsl:message terminate="no">translation.odf2oox.dateFormat</xsl:message>
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
              select="/*/office:automatic-styles/*[(self::number:date-style or self::number:time-style) and @style:name=$curStyle]"
              mode="dataStyle"/>
          </xsl:variable>
          <xsl:attribute name="w:instr">
            <xsl:value-of select="concat($fieldType,' \@ &quot;',$dataStyle,'&quot;')"/>
          </xsl:attribute>
          <w:r>
            <w:rPr>
              <xsl:call-template name="InsertLanguage"/>
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
    <xsl:value-of select="." xml:space="preserve"/>
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

  <xsl:template match="text:time|text:creation-time|text:editing-duration|text:print-time|text:modification-time" mode="paragraph">
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
              select="/*/office:automatic-styles/*[(self::number:time-style or self::number:date-style) and @style:name=$curStyle]"
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
      <!--
      makz: Commented out for bugfix 2088835
      -->
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
      <!--
      makz: Commented out for bugfix 2088835
      <xsl:apply-templates mode="paragraph"/>
      -->
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
      <xsl:apply-templates mode="paragraph"/>
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
      <xsl:apply-templates mode="paragraph"/>
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
      <xsl:apply-templates mode="paragraph"/>
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
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <!-- Templates Fields -->
  <xsl:template match="text:template-name" mode="paragraph">
    <w:fldSimple w:instr=" TEMPLATE   \* MERGEFORMAT ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:editing-cycles[not(@text:fixed='true')]" mode="paragraph">
    <w:fldSimple w:instr=" REVNUM ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <!-- chapter field. -->
  <xsl:template match="text:chapter" mode="paragraph">
    <xsl:if test="@text:outline-level">
      <!-- COMMENT : if the style changes name in the application, it may not be found and cause an error. -->
      <xsl:message terminate="no">translation.odf2oox.chapterField</xsl:message>

      <xsl:variable name="outline-level" select="@text:outline-level"/>
      <!-- find the style to match -->
      <xsl:variable name="style">
        <xsl:for-each select="document('styles.xml')">
          <xsl:choose>
            <xsl:when
              test="office:document-styles/office:styles/style:style[@style:default-outline-level=$outline-level]/@style:display-name">
              <xsl:value-of
                select="office:document-styles/office:styles/style:style[@style:default-outline-level=$outline-level]/@style:display-name"
              />
            </xsl:when>
            <xsl:when
              test="office:document-styles/office:styles/style:style[@style:default-outline-level=$outline-level and not(@style:display-name)]">
              <xsl:value-of
                select="office:document-styles/office:styles/style:style[@style:default-outline-level=$outline-level]/@style:name"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="styleName">
                <xsl:value-of
                  select="office:document-styles/office:styles/text:outline-style/text:outline-level-style[@text:level=$outline-level]/@text:style-name"
                />
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="key('styles',$styleName)/@style:display-name">
                  <xsl:value-of select="key('styles',$styleName)/@style:display-name"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$styleName"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:variable>

      <!-- if field displays number, find number associated to style. -->
      <xsl:if test="contains(@text:display, 'number')">
        <w:fldSimple>
          <xsl:attribute name="w:instr">
            <xsl:choose>
              <xsl:when test="@text:display = 'plain-number' ">
                <xsl:value-of select="concat('STYLEREF &quot;',$style,'&quot; \n \t ')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat('STYLEREF &quot;',$style,'&quot; \n ')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <w:r>
            <w:rPr>
              <w:noProof/>
            </w:rPr>
            <xsl:apply-templates mode="text"/>
          </w:r>
        </w:fldSimple>
      </xsl:if>

      <!-- if field displays name, convert into a reference to default heading style. -->
      <xsl:if test="contains(@text:display, 'name')">
        <w:fldSimple>
          <xsl:attribute name="w:instr">
            <xsl:value-of select="concat('STYLEREF &quot;',$style,'&quot; \* MERGEFORMAT')"
            />
          </xsl:attribute>
          <w:r>
            <w:rPr>
              <w:noProof/>
            </w:rPr>
            <xsl:apply-templates mode="text"/>
          </w:r>
        </w:fldSimple>
      </xsl:if>
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
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <!-- simple variables and user variables-->
  <xsl:template match="text:variable-set" mode="paragraph">
    <xsl:call-template name="InsertVariableField"/>
  </xsl:template>

  <!-- 
  Summary: inserts field declarations
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 2.11.2007
  -->
  <xsl:template name="InsertUserFieldDeclaration">
    <xsl:apply-templates select="preceding-sibling::text:user-field-decls/text:user-field-decl" mode="user-field-decl"/>
  </xsl:template>

  <!-- convert user-field-decl into regular variable field. -->
  <xsl:template match="text:user-field-decl" mode="user-field-decl">
    <xsl:call-template name="InsertVariableField"/>
  </xsl:template>

  <!--
  Summary: Templates converts fields which refers to declarations
  Author: Clever Age
  -->
  <xsl:template match="text:variable-get | text:user-field-get" mode="paragraph">
    <xsl:variable name="varName">
      <xsl:call-template name="SuppressForbiddenChars">
        <xsl:with-param name="string" select="@text:name"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="not(@text:display='none')">
      <w:fldSimple w:instr="{concat(' REF &quot;', $varName, '&quot; ')}">
        <w:r>
          <xsl:apply-templates mode="text"/>
        </w:r>
      </w:fldSimple>
    </xsl:if>
  </xsl:template>

  <xsl:template match="text:variable-input | text:user-field-input" mode="paragraph">
    <xsl:variable name="varName">
      <xsl:call-template name="SuppressForbiddenChars">
        <xsl:with-param name="string" select="@text:name"/>
      </xsl:call-template>
    </xsl:variable>
    <w:fldSimple w:instr="{concat(' ASK ', $varName, ' &quot;', @text:description, '&quot; ')}"/>
    <xsl:if test="not(@text:display='none')">
      <w:fldSimple w:instr="{concat(' REF &quot;', $varName, '&quot; ')}">
        <w:r>
          <xsl:apply-templates mode="text"/>
        </w:r>
      </w:fldSimple>
    </xsl:if>
  </xsl:template>

  <!-- report lost fields -->
  <xsl:template match="text:description" mode="paragraph">
    <xsl:message terminate="no">translation.odf2oox.descriptionField</xsl:message>
  </xsl:template>

  <xsl:template match="text:printed-by" mode="paragraph">
    <xsl:message terminate="no">translation.odf2oox.printedByField</xsl:message>
  </xsl:template>

  <xsl:template match="text:page-variable-set | text:page-variable-get" mode="paragraph">
    <xsl:message terminate="no">translation.odf2oox.pageVariableField</xsl:message>
    <xsl:apply-templates mode="paragraph"/>
  </xsl:template>

  <xsl:template match="text:dde-connection-decls[text:dde-connection-decl/@text:name]">
    <!-- lost because not in the spec, although DDE and DDEAUTO are available in Word -->
    <xsl:message terminate="no">translation.odf2oox.ddeConnection</xsl:message>
  </xsl:template>

  <xsl:template match="text:expression" mode="paragraph">
    <xsl:message terminate="no">translation.odf2oox.formulaField</xsl:message>
    <xsl:apply-templates mode="paragraph"/>
  </xsl:template>

  <!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

  <xsl:template name="InsertLanguage">
    <xsl:choose>
      <xsl:when test="$default-language">
        <w:lang w:val="{$default-language}"/>
      </xsl:when>
      <xsl:otherwise>
        <w:noProof/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
  Summary: insert declaration of variable, and potentially a reference to display it.
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 31.10.2007
  -->
  <xsl:template name="InsertVariableField">
    <xsl:variable name="varName">
      <xsl:call-template name="SuppressForbiddenChars">
        <xsl:with-param name="string" select="@text:name"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="varValue">
      <xsl:choose>
        <xsl:when test="(@office:value-type = 'float' or @office:value-type = 'percentage') and @office:value">
          <xsl:value-of select="@office:value"/>
        </xsl:when>
        <xsl:when test="@office:value-type = 'currency' and (@office:value and @office:currency)">
          <xsl:value-of select="concat(@office:value, @office:currency)"/>
        </xsl:when>
        <xsl:when test="@office:value-type = 'date' and @office:date-value">
          <xsl:value-of select="@office:date-value"/>
        </xsl:when>
        <xsl:when test="@office:value-type = 'time' and @office:time-value">
          <xsl:value-of select="@office:time-value"/>
        </xsl:when>
        <xsl:when test="@office:value-type = 'boolean' and @office:boolean-value">
          <xsl:value-of select="@office:boolean-value"/>
        </xsl:when>
        <xsl:when test="@office:value-type = 'string' and @office:string-value">
          <xsl:value-of select="@office:string-value"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="child::text()"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="bookmarkID">
      <xsl:call-template name="GenerateBookmarkId">
        <xsl:with-param name="TextName">
          <xsl:value-of select="@text:name"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:instrText xml:space="preserve"> SET <xsl:value-of select="$varName"/> "<xsl:value-of select="$varValue"/>" \* MERGEFORMAT </w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="seperate"/>
    </w:r>
    <w:bookmarkStart w:id="{bookmarkID}" w:name="{$varName}" />
    <w:r>
      <w:rPr>
        <w:noProof/>
      </w:rPr>
      <w:t>
        <xsl:value-of select="$varValue"/>
      </w:t>
    </w:r>
    <w:bookmarkEnd w:id="{bookmarkID}"/>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <xsl:if test="not(@text:display='none' or self::text:user-field-decl)">
      <w:fldSimple w:instr="{concat(' REF ', $varName, ' ')}">
        <w:r>
          <xsl:apply-templates mode="text"/>
        </w:r>
      </w:fldSimple>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

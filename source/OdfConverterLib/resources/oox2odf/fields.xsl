<?xml version="1.0" encoding="UTF-8"?>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="w number wp xlink r b">

  <!-- Date and Time Fields -->
  <xsl:template name="InsertDate">
    <xsl:param name="dateText"/>
    <text:date>
      <xsl:call-template name="InsertDateContent">
        <xsl:with-param name="dateText" select="$dateText"/>
      </xsl:call-template>
    </text:date>
  </xsl:template>

  <xsl:template name="InsertDateType">
    <xsl:param name="fieldCode"/>
    <xsl:param name="fieldType"/>
    <xsl:choose>
      <xsl:when test="$fieldType = 'DATE'">
        <xsl:call-template name="InsertDate">
          <xsl:with-param name="dateText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$fieldType = 'CREATEDATE' or contains($fieldCode, 'CreateDate') or contains($fieldCode, 'CreateTime')">
        <xsl:call-template name="InsertCreationDate">
          <xsl:with-param name="dateText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$fieldType = 'PRINTDATE' or contains($fieldCode, 'PrintDate')">
        <xsl:call-template name="InsertPrintDate">
          <xsl:with-param name="dateText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$fieldType = 'SAVEDATE' or contains($fieldCode,'LastSavedTime') or contains($fieldCode, 'SaveDate')">
        <xsl:call-template name="InsertModificationDate">
          <xsl:with-param name="dateText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertCreationDate">
    <xsl:param name="dateText"/>
    <text:creation-date>
      <xsl:call-template name="InsertDateContent">
        <xsl:with-param name="dateText" select="$dateText"/>
      </xsl:call-template>
    </text:creation-date>
  </xsl:template>

  <xsl:template name="InsertPrintDate">
    <xsl:param name="dateText"/>
    <text:print-date>
      <xsl:call-template name="InsertDateContent">
        <xsl:with-param name="dateText" select="$dateText"/>
      </xsl:call-template>
    </text:print-date>
  </xsl:template>

  <xsl:template name="InsertModificationDate">
    <xsl:param name="dateText"/>
    <text:modification-date>
      <xsl:call-template name="InsertDateContent">
        <xsl:with-param name="dateText" select="$dateText"/>
      </xsl:call-template>
    </text:modification-date>
  </xsl:template>

  <xsl:template name="InsertDateContent">
    <xsl:param name="dateText"/>
    <xsl:attribute name="style:data-style-name">
      <xsl:value-of select="generate-id(.)"/>
    </xsl:attribute>
    <xsl:attribute name="text:date-value">
      <xsl:value-of select="$dateText"/>
    </xsl:attribute>
    <xsl:value-of select="$dateText"/>
  </xsl:template>

  <xsl:template name="InsertTimeType">
    <xsl:param name="fieldCode"/>
    <xsl:param name="fieldType"/>

    <xsl:choose>
      <xsl:when test="$fieldType = 'TIME' ">
        <xsl:call-template name="InsertTime">
          <xsl:with-param name="timeText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
      <!-- EDITTIME, DOCPROPERTY TotalEditingTime-->
      <xsl:when
        test="$fieldType = 'EDITTIME' or contains($fieldCode,'TotalEditingTime') or contains($fieldCode, 'EditTime')">
        <xsl:call-template name="InsertEditTime">
          <xsl:with-param name="timeText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
      <!-- DOCPROPERTY CreateTime-->
      <xsl:when test="contains($fieldCode,'CreateTime') ">
        <xsl:call-template name="InsertCreationTime">
          <xsl:with-param name="timeText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertCreationTime">
    <xsl:param name="timeText"/>
    <text:creation-time>
      <xsl:attribute name="style:data-style-name">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
      <xsl:attribute name="text:time-value">
        <xsl:value-of select="$timeText"/>
      </xsl:attribute>
    </text:creation-time>
  </xsl:template>

  <xsl:template name="InsertTime">
    <xsl:param name="timeText"/>
    <text:time>
      <xsl:attribute name="style:data-style-name">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
      <xsl:attribute name="text:time-value">
        <xsl:value-of select="$timeText"/>
      </xsl:attribute>
    </text:time>
  </xsl:template>

  <xsl:template name="InsertEditTime">
    <xsl:param name="timeText"/>
    <text:editing-duration>
      <xsl:attribute name="style:data-style-name">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
    </text:editing-duration>
  </xsl:template>

  <!-- process a field code -->
  <xsl:template match="w:instrText">
    <xsl:param name="parentRunNode"/>
    <text:span text:style-name="{generate-id(parent::w:r)}">
      <!-- rebuild the field code using a series of instrText, in current run or followings -->
      <xsl:variable name="fieldCode">
        <xsl:call-template name="BuildFieldCode"/>
      </xsl:variable>
      <!-- first field instruction. Should contains field type. If not, switch to next instruction -->
      <xsl:variable name="fieldType">
        <xsl:call-template name="GetFieldTypeFromCode">
          <xsl:with-param name="fieldCode" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$fieldType = 'XE' or $fieldType = 'xe' ">
          <xsl:call-template name="InsertIndexMark">
            <xsl:with-param name="instrText" select="$fieldCode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'SET' or $fieldType = 'set' ">
          <xsl:call-template name="InsertUserVariable">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'HYPERLINK' or $fieldType = 'hyperlink' ">
          <xsl:call-template name="InsertHyperlinkField">
            <xsl:with-param name="parentRunNode" select="$parentRunNode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'REF' or $fieldType = 'ref' ">
          <xsl:call-template name="InsertCrossReference">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
          </xsl:call-template>
        </xsl:when>
        <!--  possible date types: DATE, PRINTDATE, SAVEDATE, CREATEDATE, INFO CreateDate, INFO PrintDate, INFO Savedate-->
        <xsl:when
          test="contains($fieldCode, 'DATE' ) or contains($fieldCode,  'date') 
          or contains($fieldCode,'LastSavedTime') or contains($fieldCode,'CreateDate')
          or contains($fieldCode, 'PrintDate') or contains($fieldCode,'SaveDate') or contains($fieldCode,'CreateTime')">
          <xsl:call-template name="InsertDateType">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
            <xsl:with-param name="fieldType" select="$fieldType"/>
          </xsl:call-template>
        </xsl:when>
        <!--page-count NUMPAGE, DOCPROPERTY Pages-->
        <xsl:when
          test="$fieldType = 'NUMPAGE' or  contains(.,'NUMPAGES') or $fieldType = 'numpage' or contains($fieldCode,'Pages')">
          <xsl:call-template name="InsertPageCount"/>
        </xsl:when>
        <xsl:when test="$fieldType = 'PAGE' or $fieldType = 'page' ">
          <xsl:call-template name="InsertPageNumber"/>
        </xsl:when>
        <!-- possible time types: TIME, EDITTIME, DOCPROPERTY CreateTime, DOCPROPERTY TotalEditingTime,  INFO EditTime-->
        <xsl:when
          test="$fieldType = 'TIME' or $fieldType = 'time' or contains($fieldCode,'TotalEditingTime') or contains($fieldCode, 'EditTime')">
          <xsl:call-template name="InsertTimeType">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
            <xsl:with-param name="fieldType" select="$fieldType"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'USERINITIALS' or $fieldType = 'userinitials' ">
          <xsl:call-template name="InsertUserInitials"/>
        </xsl:when>
        <xsl:when test="$fieldType = 'USERNAME' or $fieldType = 'username' ">
          <xsl:call-template name="InsertUserName"/>
        </xsl:when>
        <!--initiial creator name   AUTHOR and DOCPROPERTY Author-->
        <xsl:when
          test="$fieldType = 'AUTHOR' or $fieldType = 'author'  or contains($fieldCode,'Author')">
          <xsl:call-template name="InsertAuthor"/>
        </xsl:when>
        <!--caption field  from which Index of Figures is created -->
        <xsl:when test="$fieldType = 'SEQ' or $fieldType = 'seq' ">
          <xsl:call-template name="InsertSequence">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
            <xsl:with-param name="sequenceContext" select="following::w:r[w:t][1]"/>
          </xsl:call-template>
        </xsl:when>
        <!--creator name LASTSAVEDBY, DOCPROPERTY LastSavedBy-->
        <xsl:when
          test="$fieldType = 'LASTSAVEDBY' or $fieldType = 'lastsavedby'  or  contains($fieldCode,'LastSavedBy')">
          <xsl:call-template name="InsertCreator"/>
        </xsl:when>
        <!--editing cycles number REVNUM, DOCPROPERTY RevisionNumber, INFO RevNum-->
        <xsl:when
          test="$fieldType = 'REVNUM' or $fieldType = 'revnum' 
          or contains($fieldCode,'RevisionNumber') or contains($fieldCode, 'RevNum')">
          <xsl:call-template name="InsertEditingCycles"/>
        </xsl:when>
        <!--FILENAME, INFO FileName-->
        <xsl:when
          test="$fieldType = 'FILENAME' or $fieldType = 'filename' or contains($fieldCode,'FileName')">
          <xsl:call-template name="InsertFileName"/>
        </xsl:when>
        <!-- KEYWORDS, DOCPROPERTY Keywords -->
        <xsl:when
          test="contains($fieldCode,'KEYWORDS') or contains($fieldCode,'keywords') or contains($fieldCode,'Keywords')">
          <xsl:call-template name="InsertKeywords"/>
        </xsl:when>
        <!-- DOCPROPERTY Company, INFO Company-->
        <xsl:when test="contains($fieldCode,'Company') or contains($fieldCode, 'company')">
          <xsl:call-template name="InsertCompany"/>
        </xsl:when>
        <xsl:when test="$fieldType = 'USERADDRESS' or $fieldType = 'useraddress' ">
          <xsl:call-template name="InsertUserAddress"/>
        </xsl:when>
        <!--TEMPLATE, DOCPROPERTY Template-->
        <xsl:when
          test="$fieldType = 'TEMPLATE' or $fieldType = 'template'  or contains($fieldCode,'Template')">
          <xsl:call-template name="InsertTemplate"/>
        </xsl:when>
        <!--NUMWORDS, DOCPROPERTY Words-->
        <xsl:when
          test="$fieldType = 'NUMWORDS' or $fieldType = 'numwords' or contains($fieldCode,'Words')">
          <xsl:call-template name="InsertWordCount"/>
        </xsl:when>
        <!--NUMCHARS and DOCPROPERTY Characters,  INFO NumChars-->
        <xsl:when
          test="$fieldType = 'NUMCHARS' or $fieldType = 'numchars' 
          or contains($fieldCode,'Characters') or contains($fieldCode, 'NumChars')">
          <xsl:call-template name="InsertCharacterCount"/>
        </xsl:when>
        <!-- DOCPROPERTY Paragraphs, INFO Paragraphs-->
        <xsl:when test="contains($fieldCode,'Paragraphs') or contains($fieldCode,'paragraphs') ">
          <xsl:call-template name="InsertCompany"/>
        </xsl:when>
        <!-- COMMENTS, DOCPROPERTY Comments-->
        <xsl:when test="contains($fieldCode,'COMMENTS') or contains($fieldCode,'Comments') ">
          <xsl:call-template name="InsertComments"/>
        </xsl:when>
        <!--document subject SUBJECT, DOCPROPERTY Subject-->
        <xsl:when
          test="contains($fieldCode,'SUBJECT') or contains($fieldCode,'Subject') or contains($fieldCode,'subject')">
          <xsl:call-template name="InsertSubject"/>
        </xsl:when>
        <!--document title TITLE, DOCPROPERTY Title-->
        <xsl:when
          test="contains($fieldCode,'TITLE') or contains($fieldCode,'Title') or contains($fieldCode,'title')">
          <xsl:call-template name="InsertSubject"/>
        </xsl:when>
        <!--bibliography citation-->
        <xsl:when test="$fieldType = 'CITATION' or $fieldType = 'citation' ">
          <xsl:call-template name="InsertTextBibliographyMark">
            <xsl:with-param name="TextIdentifier">
              <xsl:value-of
                select="substring-before(substring-after(self::node(), 'CITATION '), ' \')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
      </xsl:choose>
    </text:span>
  </xsl:template>

  <xsl:template match="w:instrText" mode="automaticstyles">
    <!-- rebuild the field code using a series of instrText, in current run or followings -->
    <xsl:variable name="fieldCode">
      <xsl:call-template name="BuildFieldCode"/>
    </xsl:variable>
    <!-- first field instruction. Should contains field type. If not, switch to next instruction -->
    <xsl:variable name="fieldType">
      <xsl:call-template name="GetFieldTypeFromCode">
        <xsl:with-param name="fieldCode" select="$fieldCode"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <!--  possible date types: DATE, PRINTDATE, SAVEDATE, CREATEDATE-->
      <xsl:when test="contains($fieldType,'DATE')">
        <xsl:call-template name="InsertDateStyle">
          <xsl:with-param name="dateText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
      <!-- possible time types: TIME, EDITTIME-->
      <xsl:when test="contains($fieldType,'TIME')">
        <xsl:call-template name="InsertTimeStyle">
          <xsl:with-param name="timeText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--cross-reference-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'REF')]" mode="fields">
    <xsl:call-template name="InsertCrossReference">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <!-- build a field code using current instrText and recursively forward (go to next instrText) -->
  <xsl:template name="BuildFieldCode">
    <xsl:param name="instrText" select="."/>
    <xsl:param name="fieldCode"/>
    <!-- context must be w:instrText -->
    <xsl:choose>
      <!-- if next sibling instrText -->
      <xsl:when test="$instrText/following-sibling::w:instrText">
        <xsl:call-template name="BuildFieldCode">
          <xsl:with-param name="instrText" select="$instrText/following-sibling::w:instrText[1]"/>
          <xsl:with-param name="fieldCode" select="concat($fieldCode, $instrText/text())"/>
        </xsl:call-template>
      </xsl:when>
      <!-- if next run with instrText, before end of field -->
      <xsl:when
        test="$instrText/parent::w:r/following-sibling::w:r/*[self::w:instrText or self::w:fldChar[@w:fldCharType = 'end']][1][self::w:instrText]">
        <xsl:call-template name="BuildFieldCode">
          <!-- we know now that first run having instrText is before end of field -->
          <xsl:with-param name="instrText"
            select="$instrText/parent::w:r/following-sibling::w:r[w:instrText][1]/w:instrText[1]"/>
          <xsl:with-param name="fieldCode" select="concat($fieldCode, $instrText/text())"/>
        </xsl:call-template>
      </xsl:when>
      <!-- if next paragraph with instrText, before end of field :
        Find first paragraph having instrText before end of field, and then first run having instrText before end of field -->
      <xsl:when
        test="$instrText/ancestor::w:p/following-sibling::w:p/w:r[w:instrText or w:fldChar[@w:fldCharType = 'end']][1]/*[self::w:instrText or self::w:fldChar[@w:fldCharType = 'end']][1][self::w:instrText]">
        <!-- check first is field does not end in context paragraph -->
        <xsl:choose>
          <xsl:when
            test="not($instrText/parent::w:r/following-sibling::w:r/w:fldChar[@w:fldCharType = 'end'])">
            <xsl:call-template name="BuildFieldCode">
              <!-- we know now that first run having instrText is before end of field -->
              <xsl:with-param name="instrText"
                select="$instrText/ancestor::w:p/following-sibling::w:p[w:r/w:instrText][1]/w:r[w:instrText][1]/w:instrText[1]"/>
              <xsl:with-param name="fieldCode" select="concat($fieldCode, $instrText/text())"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat($fieldCode, $instrText/text())"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat($fieldCode, $instrText/text())"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- return a field type if it is contained in param string -->
  <xsl:template name="GetFieldTypeFromCode">
    <xsl:param name="fieldCode"/>
    <!-- field can start with space, but first none-space text is field code -->
    <xsl:variable name="newFieldCode">
      <xsl:call-template name="suppressFieldCodeFirstSpaceChar">
        <xsl:with-param name="string" select="$fieldCode"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="contains($newFieldCode, ' ')">
        <xsl:value-of select="substring-before($newFieldCode, ' ')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$newFieldCode"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- suppress first space char in a given string (used for field codes -->
  <xsl:template name="suppressFieldCodeFirstSpaceChar">
    <xsl:param name="string"/>
    <xsl:choose>
      <xsl:when test="substring($string, 1, 1) = ' ' ">
        <xsl:call-template name="suppressFieldCodeFirstSpaceChar">
          <xsl:with-param name="string" select="substring($string, 2, string-length($string) - 1)"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$string"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- user field declaration -->
  <xsl:template name="InsertUserVariable">
    <xsl:param name="fieldCode"/>
    <!-- troncate field to find arguments -->
    <xsl:variable name="fieldName">
      <xsl:variable name="newFieldCode">
        <xsl:call-template name="suppressFieldCodeFirstSpaceChar">
          <xsl:with-param name="string" select="substring-after($fieldCode, 'SET')"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="contains($newFieldCode, ' ')">
          <xsl:value-of select="substring-before($newFieldCode, ' ')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$newFieldCode"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="fieldValue">
      <xsl:variable name="newFieldCode">
        <xsl:call-template name="suppressFieldCodeFirstSpaceChar">
          <xsl:with-param name="string"
            select="substring-after(substring-after($fieldCode, 'SET'), $fieldName)"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="substring($newFieldCode, 1,1) = &apos;&quot;&apos;">
          <xsl:value-of
            select="substring-before(substring-after($newFieldCode, &apos;&quot;&apos;), &apos;&quot;&apos;)"
          />
        </xsl:when>
        <xsl:when test="contains($newFieldCode, ' ')">
          <xsl:value-of select="substring-before($newFieldCode, ' ')"/>
        </xsl:when>
        <!--xsl:when test="$newFieldCode = '' ">
          < at least a blank space >
          <xsl:text xml:space="preserve"> </xsl:text>
          </xsl:when-->
        <xsl:otherwise>
          <xsl:value-of select="$newFieldCode"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- COMMENT : variable fields should be declared before set, but application should support direct set before declaration,
      and it is too complex to find wether the field has already been declared. So no declaration is performed. -->
    <text:variable-set>
      <xsl:attribute name="text:name">
        <xsl:value-of select="$fieldName"/>
      </xsl:attribute>
      <xsl:variable name="valueType">
        <xsl:choose>
          <xsl:when test="number($fieldValue)">float</xsl:when>
          <xsl:otherwise>string</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="office:value-type">
        <xsl:value-of select="$valueType"/>
      </xsl:attribute>
      <!-- TODO : find the best matching type for the value -->
      <xsl:choose>
        <xsl:when test="$valueType = 'float' or $valueType = 'percentage' ">
          <xsl:attribute name="office:value">
            <xsl:value-of select="$fieldValue"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$valueType = 'currency' ">
          <xsl:attribute name="office:value">
            <xsl:value-of select="$fieldValue"/>
          </xsl:attribute>
          <xsl:attribute name="office:currency"/>
        </xsl:when>
        <xsl:when test="$valueType = 'date' ">
          <xsl:attribute name="office:date-value">
            <xsl:value-of select="$fieldValue"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$valueType = 'time' ">
          <xsl:attribute name="office:time-value">
            <xsl:value-of select="$fieldValue"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$valueType = 'boolean' ">
          <xsl:attribute name="office:boolean-value">
            <xsl:value-of select="$fieldValue"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$valueType = 'string' ">
          <xsl:attribute name="office:string-value">
            <xsl:value-of select="$fieldValue"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
      <xsl:attribute name="text:display">none</xsl:attribute>
    </text:variable-set>
  </xsl:template>

  <!-- alphabetical index mark -->
  <xsl:template name="InsertIndexMark">
    <xsl:param name="instrText"/>
    <xsl:variable name="Value">
      <xsl:value-of select="substring-before(substring-after($instrText,'&quot;'),'&quot;')"
      />
    </xsl:variable>

    <text:alphabetical-index-mark>
      <xsl:choose>
        <xsl:when test="not(contains($Value, ':'))">
          <xsl:attribute name="text:string-value">
            <xsl:value-of select="$Value"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($Value, ':')">
          <xsl:variable name="TextKey1">
            <xsl:value-of select="substring-before($Value, ':')"/>
          </xsl:variable>
          <xsl:variable name="TextKey2">
            <xsl:value-of select="substring-after($Value, ':')"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="contains($TextKey2, ':')">
              <xsl:attribute name="text:string-value">
                <xsl:value-of select="substring-after($TextKey2, ':')"/>
              </xsl:attribute>
              <xsl:attribute name="text:key2">
                <xsl:value-of select="substring-before($TextKey2, ':')"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="text:string-value">
                <xsl:value-of select="$TextKey2"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:attribute name="text:key1">
            <xsl:value-of select="$TextKey1"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </text:alphabetical-index-mark>
  </xsl:template>

  <xsl:template name="InsertCrossReference">
    <xsl:param name="fieldCode"/>
    <text:bookmark-ref text:reference-format="text">
      <xsl:attribute name="text:ref-name">
        <xsl:value-of select="substring-before(substring-after($fieldCode,'REF '),' \')"/>
      </xsl:attribute>
      <xsl:apply-templates select="descendant::w:r/child::node()"/>
    </text:bookmark-ref>
  </xsl:template>

  <xsl:template name="InsertFieldCharFieldContent">
    <xsl:variable name="fieldCharId" select="generate-id(.)"/>
    <!-- element that starts field-->
    <xsl:variable name="fieldBegin"
      select="preceding::w:fldChar[@w:fldCharType='begin' and not(following::w:fldChar
      [@w:fldCharType='end' and generate-id(following::w:instrText[contains(self::node(),'REF')][1]) = $fieldCharId ]) ][1] "/>
    <!-- element that ends field-->
    <xsl:variable name="fieldEnd"
      select="following::w:fldChar[@w:fldCharType='end' and not(preceding::w:fldChar
      [@w:fldCharType='begin' and generate-id(preceding::w:instrText[contains(self::node(),'REF')][1]) = fieldCharId ]) ][1] "/>
    <!--paragraph that contains field-->
    <xsl:variable name="fieldParagraph" select="generate-id(ancestor::w:p)"/>

    <!--  inserts field  text contents from beginning to ending  element -->
    <xsl:for-each
      select="$fieldBegin/following::w:t[generate-id(ancestor::w:p) = $fieldParagraph
      and generate-id(preceding::w:fldChar[1]) != generate-id($fieldEnd) ]  ">
      <xsl:apply-templates/>
    </xsl:for-each>
  </xsl:template>

  <!--document title TITLE, DOCPROPERTY Title-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'TITLE') or contains(@w:instr, 'Title')]"
    mode="fields">
    <text:title>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:title>
  </xsl:template>

  <xsl:template match="w:fldSimple[contains(@w:instr,'USERNAME')]" mode="fields">
    <xsl:call-template name="InsertUserName"/>
  </xsl:template>

  <xsl:template name="InsertUserName">
    <text:author-name text:fixed="false">
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:author-name>
  </xsl:template>

  <!--AUTHOR and DOCPROPERTY Author-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'AUTHOR') or contains(@w:instr,' Author')]"
    mode="fields">
    <xsl:call-template name="InsertAuthor"/>
  </xsl:template>

  <xsl:template name="InsertAuthor">
    <text:initial-creator>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:initial-creator>
  </xsl:template>

  <!--creator-name LASTSAVEDBY, DOCPROPERTY LastSavedBy-->
  <xsl:template
    match="w:fldSimple[contains(@w:instr,'LASTSAVEDBY') or contains(@w:instr,'LastSavedBy')]"
    mode="fields">
    <xsl:call-template name="InsertCreator"/>
  </xsl:template>

  <xsl:template name="InsertCreator">
    <text:creator>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:creator>
  </xsl:template>

  <!--user initials-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'USERINITIALS')]" mode="fields">
    <xsl:call-template name="InsertUserInitials"/>
  </xsl:template>

  <!--user initials-->
  <xsl:template name="InsertUserInitials">
    <text:author-initials>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:author-initials>
  </xsl:template>

  <!--chapter name or chapter number-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'STYLEREF')]" mode="fields">
    <text:chapter>
      <xsl:choose>
        <xsl:when test="self::node()[contains(@w:instr,'\n')]">
          <xsl:attribute name="text:display">number</xsl:attribute>
        </xsl:when>
        <xsl:when test="self::node()[contains(@w:instr,'\*')]">
          <xsl:attribute name="text:display">name</xsl:attribute>
        </xsl:when>
      </xsl:choose>
      <xsl:if test="self::node()[contains(@w:instr,'Heading')]">
        <xsl:attribute name="text:outline-level">
          <xsl:value-of
            select="substring(substring-before(@w:instr,' \'),string-length(substring-before(@w:instr,' \')))"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:chapter>
  </xsl:template>

  <!--document subject SUBJECT, DOCPROPERTY Subject-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'SUBJECT') or contains(@w:instr, 'Subject')]"
    mode="fields">
    <xsl:call-template name="InsertSubject"/>
  </xsl:template>

  <xsl:template name="InsertSubject">
    <text:subject>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:subject>
  </xsl:template>

  <xsl:template match="w:fldSimple">
    <text:span text:style-name="{generate-id(w:r)}">
      <xsl:apply-templates select="." mode="fields"/>
    </text:span>
  </xsl:template>

  <!--  possible date types: DATE, PRINTDATE, SAVEDATE, CREATEDATE, 
    DOPCPROPERTY LastSavedTime, DOCPROPERTY CreateTime, INFO CreateDate, INFO PrintDate, INFO Savedate-->
  <xsl:template
    match="w:fldSimple[contains(@w:instr,'DATE') or contains(@w:instr,'LastSavedTime') or contains(@w:instr, 'CreateTime')
    or contains(@w:instr,'CreateDate') or contains(@w:instr, 'PrintDate') or contains(@w:instr, 'SaveDate')] "
    mode="fields">
    <xsl:variable name="fieldType">
      <xsl:call-template name="GetFieldTypeFromCode">
        <xsl:with-param name="fieldCode" select="@w:instr"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:call-template name="InsertDateType">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
      <xsl:with-param name="fieldType" select="$fieldType"/>
    </xsl:call-template>
  </xsl:template>

  <!-- possible time types: TIME, EDITTIME, DOCPROPERTY TotalEditingTime, INFO EditTime-->
  <xsl:template
    match="w:fldSimple[contains(@w:instr,'TIME') or contains(@w:instr, 'TotalEditingTime') or contains(@w:instr, 'EditTime')]"
    mode="fields">
    <xsl:variable name="fieldType">
      <xsl:call-template name="GetFieldTypeFromCode">
        <xsl:with-param name="fieldCode" select="@w:instr"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:call-template name="InsertTimeType">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
      <xsl:with-param name="fieldType" select="$fieldType"/>
    </xsl:call-template>
  </xsl:template>

  <!--page number-->
  <xsl:template
    match="w:fldSimple[(contains(@w:instr,'PAGE ') or contains(@w:instr,'page ')) and not(contains(@w:instr,'NUMPAGE'))]"
    mode="fields">
    <xsl:call-template name="InsertPageNumber">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <!--page-count NUMPAGE, DOCPROPERTY Pages-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'NUMPAGE') or contains(@w:instr,'Pages')]"
    mode="fields">
    <xsl:call-template name="InsertPageCount">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="InsertPageCount">
    <text:page-count/>
  </xsl:template>


  <xsl:template
    match="w:fldSimple[contains(@w:instr,'DATE') or contains(@w:instr,'LastSavedTime') or contains(@w:instr, 'CreateTime')
    or contains(@w:instr,'CreateDate') or contains(@w:instr, 'PrintDate') or contains(@w:instr, 'SaveDate')]"
    mode="automaticstyles">
    <xsl:call-template name="InsertDateStyle">
      <xsl:with-param name="dateText" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <!--caption field  from which Index of Figures is created -->
  <xsl:template match="w:fldSimple[contains(@w:instr,'SEQ')]" mode="fields">
    <xsl:call-template name="InsertSequence">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
      <xsl:with-param name="sequenceContext" select="w:r"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="InsertSequence">
    <xsl:param name="fieldCode"/>
    <xsl:param name="sequenceContext"/>
    <xsl:variable name="refType">
      <xsl:value-of select="substring-before(substring-after($fieldCode,'SEQ '),' ')"/>
    </xsl:variable>

    <text:sequence text:ref-name="{concat('ref',concat($refType,$sequenceContext/w:t))}"
      text:name="{$refType}" text:formula="{concat(concat('ooow:',$refType),'+1')}">
      <xsl:apply-templates select="$sequenceContext/child::node()"/>
    </text:sequence>
  </xsl:template>

  <xsl:template name="InsertDateStyle">
    <xsl:param name="dateText"/>
    <xsl:variable name="FormatDate">
      <xsl:value-of
        select="substring-before(substring-after($dateText, '&quot;'), '&quot;')"/>
    </xsl:variable>
    <!-- some of the DOCPROPERTY date field types have constant date format, 
      which is not saved in fieldCode so it need to be given directly in these cases-->
    <xsl:choose>
      <xsl:when test="contains($dateText, 'CreateTime') or contains($dateText,'LastSavedTime')">
        <xsl:call-template name="InsertDocprTimeStyle"/>
      </xsl:when>
      <xsl:when
        test="contains($dateText,'CreateDate') or contains($dateText, 'SaveDate') or contains($dateText, 'PrintDate')">
        <xsl:call-template name="InsertDocprLongDateStyle"/>
      </xsl:when>
      <!--default scenario-->
      <xsl:otherwise>
        <xsl:call-template name="InsertDateFormat">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="$FormatDate"/>
          </xsl:with-param>
          <xsl:with-param name="ParamField">
            <xsl:text>DATE</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--DOCPROPERTY CreateDate, SaveDate, PrintDate default format-->
  <xsl:template name="InsertDocprLongDateStyle">
    <number:date-style style:name="{generate-id()}">
      <number:month number:style="long"/>
      <number:text>-</number:text>
      <number:day number:style="long"/>
      <number:text>-</number:text>
      <number:year number:style="long"/>
      <number:text>
        <xsl:text> </xsl:text>
      </number:text>
      <number:hours number:style="long"/>
      <number:text>:</number:text>
      <number:minutes number:style="long"/>
      <number:text>:</number:text>
      <number:seconds number:style="long"/>
      <number:text>
        <xsl:text> </xsl:text>
      </number:text>
      <number:am-pm/>
    </number:date-style>
  </xsl:template>

  <!--DOCPROPERTY CreateTime, LastSavedTime default format-->
  <xsl:template name="InsertDocprTimeStyle">
    <number:date-style style:name="{generate-id()}">
      <number:year number:style="long"/>
      <number:text>-</number:text>
      <number:month number:style="long"/>
      <number:text>-</number:text>
      <number:day number:style="long"/>
      <number:text>
        <xsl:text> </xsl:text>
      </number:text>
      <number:hours number:style="long"/>
      <number:text>:</number:text>
      <number:minutes number:style="long"/>
    </number:date-style>
  </xsl:template>

  <xsl:template
    match="w:fldSimple[contains(@w:instr,'TIME') or contains(@w:instr, 'TotalEditingTime') or contains(@w:instr, 'EditTime')]"
    mode="automaticstyles">
    <xsl:call-template name="InsertTimeStyle">
      <xsl:with-param name="timeText" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="InsertTimeStyle">
    <xsl:param name="timeText"/>
    <xsl:variable name="FormatDate">
      <xsl:value-of
        select="substring-before(substring-after($timeText, '&quot;'), '&quot;')"/>
    </xsl:variable>
    <xsl:call-template name="InsertDateFormat">
      <xsl:with-param name="FormatDate">
        <xsl:value-of select="$FormatDate"/>
      </xsl:with-param>
      <xsl:with-param name="ParamField">
        <xsl:text>TIME</xsl:text>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:fldSimple" mode="automaticstyles">
    <xsl:apply-templates select="w:r/w:rPr" mode="automaticstyles"/>
  </xsl:template>

  <xsl:template name="InsertFieldProperties">
    <xsl:param name="fieldCodeContainer" select="ancestor::w:fldSimple | ancestor::w:r/w:instrText"/>

    <xsl:variable name="fieldCode">
      <xsl:choose>
        <xsl:when test="$fieldCodeContainer/@w:instr">
          <xsl:value-of select="$fieldCodeContainer/@w:instr"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="BuildFieldCode">
            <xsl:with-param name="instrText" select="$fieldCodeContainer"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="contains($fieldCode,'Upper')">
        <xsl:attribute name="fo:text-transform">
          <xsl:text>uppercase</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="contains($fieldCode,'Lower')">
        <xsl:attribute name="fo:text-transform">
          <xsl:text>lowercase</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="contains($fieldCode,'FirstCap') or contains($fieldCode,'Caps')">
        <xsl:attribute name="fo:text-transform">
          <xsl:text>capitalize</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="contains($fieldCode,'SBCHAR')">
        <xsl:attribute name="fo:letter-spacing">
          <xsl:text>-0.018cm</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="contains($fieldCode,'DBCHAR')">
        <xsl:attribute name="fo:letter-spacing">
          <xsl:text>0.176cm</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertDateFormat">
    <xsl:param name="FormatDate"/>
    <xsl:param name="ParamField"/>
    <number:date-style>
      <xsl:attribute name="style:name">
        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>
      <xsl:call-template name="InsertFormatDateStyle">
        <xsl:with-param name="FormatDate">
          <xsl:choose>
            <xsl:when test="$FormatDate!=''">
              <xsl:value-of select="$FormatDate"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$ParamField='TIME'">
                  <xsl:text>HH:mm</xsl:text>
                </xsl:when>
                <xsl:when test="$ParamField='DATE'">
                  <xsl:text>yyyy-MM-dd</xsl:text>
                </xsl:when>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </number:date-style>
  </xsl:template>

  <xsl:template name="InsertFormatDateStyle">
    <xsl:param name="FormatDate"/>
    <xsl:param name="DateText"/>
    <xsl:choose>

      <xsl:when test="starts-with($FormatDate, 'yyyy')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:year number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'yyyy')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'yy')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:year/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'yy')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'y')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:year/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'y')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'YYYY')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:year number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'YYYY')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'YY')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:year/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'YY')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'Y')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:year/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'Y')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'MMMM')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:month number:style="long" number:textual="true"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'MMMM')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'MMM')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:month number:textual="true"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'MMM')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'MM')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:month number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'MM')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'M')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:month/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'M')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'dddd')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day-of-week number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'dddd')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'ddd')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day-of-week/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'ddd')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'dd')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'dd')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'd')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'd')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'DDDD')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day-of-week number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'DDDD')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'DDD')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day-of-week/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'DDD')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'DD')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'DD')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'D')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:day number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'D')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>


      <xsl:when test="starts-with($FormatDate, 'hh')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:hours number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'hh')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'h')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:hours/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'h')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'HH')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:hours number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'HH')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'H')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:hours/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'H')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'mm')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:minutes number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'mm')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'm')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:minutes/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'm')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'ss')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:seconds number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'ss')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 's')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:seconds/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 's')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'SS')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:seconds number:style="long"/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'SS')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'S')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:seconds/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'S')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="starts-with($FormatDate, 'am/pm')">
        <xsl:if test="$DateText">
          <number:text>
            <xsl:value-of select="$DateText"/>
          </number:text>
        </xsl:if>
        <number:am-pm/>
        <xsl:call-template name="InsertFormatDateStyle">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="substring-after($FormatDate, 'am/pm')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:otherwise>
        <xsl:variable name="Apostrof">
          <xsl:text>&apos;</xsl:text>
        </xsl:variable>

        <xsl:choose>
          <xsl:when test="contains(substring($FormatDate, 1, 1), $Apostrof)">
            <xsl:call-template name="InsertFormatDateStyle">
              <xsl:with-param name="FormatDate">
                <xsl:value-of
                  select="substring-after(substring-after($FormatDate, $Apostrof), $Apostrof)"/>
              </xsl:with-param>
              <xsl:with-param name="DateText">
                <xsl:value-of
                  select="concat($DateText, substring-before(substring-after($FormatDate, $Apostrof), $Apostrof))"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>

          <xsl:when test="substring-after($FormatDate, substring($FormatDate, 1, 1))">
            <xsl:call-template name="InsertFormatDateStyle">
              <xsl:with-param name="FormatDate">
                <xsl:value-of select="substring-after($FormatDate, substring($FormatDate, 1, 1))"/>
              </xsl:with-param>
              <xsl:with-param name="DateText">
                <xsl:value-of select="concat($DateText, substring($FormatDate, 1, 1))"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>

          <xsl:otherwise>
            <xsl:if test="$DateText">
              <number:text>
                <xsl:value-of select="$DateText"/>
              </number:text>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:sdt/w:sdtContent/w:fldSimple">
    <xsl:apply-templates/>
  </xsl:template>

  <!-- Page Number Field -->
  <xsl:template name="InsertPageNumber">
    <xsl:variable name="docName">
      <xsl:call-template name="GetDocumentName">
        <xsl:with-param name="rootId">
          <xsl:value-of select="generate-id(/node())"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$docName = 'document.xml'">
        <xsl:if test="following::w:sectPr[1]/w:pgNumType/@w:chapStyle">
          <text:chapter>
            <xsl:attribute name="text:display">
              <xsl:text>number</xsl:text>
            </xsl:attribute>
            <xsl:attribute name="text:outline-level">
              <xsl:value-of select="following::w:sectPr[1]/w:pgNumType/@w:chapStyle"/>
            </xsl:attribute>
          </text:chapter>
          <xsl:choose>
            <xsl:when test="following::w:sectPr[1]/w:pgNumType/@w:chapSep = 'period'">
              <xsl:text>.</xsl:text>
            </xsl:when>
            <xsl:when test="following::w:sectPr[1]/w:pgNumType/@w:chapSep = 'colon'">
              <xsl:text>:</xsl:text>
            </xsl:when>
            <xsl:otherwise>-</xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>

        <xsl:variable name="rId">
          <xsl:value-of
            select="document('word/_rels/document.xml.rels')/Relationships/Relationship[@Target = $docName]/@Id"
          />
        </xsl:variable>
        <xsl:for-each select="document('word/document.xml')">
          <xsl:for-each
            select="key('sectPr', '')[w:headerReference/@r:id = $rId or w:footerReference/@r:id = $rId]">
            <xsl:if test="w:pgNumType/@w:chapStyle">
              <text:chapter>
                <xsl:attribute name="text:display">
                  <xsl:text>number</xsl:text>
                </xsl:attribute>
                <xsl:attribute name="text:outline-level">
                  <xsl:value-of select="w:pgNumType/@w:chapStyle"/>
                </xsl:attribute>
              </text:chapter>
              <xsl:choose>
                <xsl:when test="w:pgNumType/@w:chapSep = 'period'">
                  <xsl:text>.</xsl:text>
                </xsl:when>
                <xsl:when test="following::w:sectPr[1]/w:pgNumType/@w:chapSep = 'colon'">
                  <xsl:text>:</xsl:text>
                </xsl:when>
                <xsl:otherwise>-</xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:variable name="WInstr">
      <xsl:value-of select="./@w:instr"/>
    </xsl:variable>
    <text:page-number>
      <xsl:attribute name="text:select-page">
        <xsl:text>current</xsl:text>
      </xsl:attribute>
      <xsl:attribute name="style:num-format">
        <xsl:choose>
          <xsl:when test="contains($WInstr, 'Arabic')">1</xsl:when>
          <xsl:when test="contains($WInstr, 'alphabetic')">a</xsl:when>
          <xsl:when test="contains($WInstr, 'ALPHABETIC')">A</xsl:when>
          <xsl:when test="contains($WInstr, 'roman')">i</xsl:when>
          <xsl:when test="contains($WInstr, 'ROMAN')">I</xsl:when>
          <xsl:otherwise>1</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:value-of select="$WInstr"/>
    </text:page-number>
  </xsl:template>

  <!-- Insert Citations -->
  <xsl:template name="InsertTextBibliographyMark">
    <xsl:param name="TextIdentifier"/>

    <xsl:variable name="Path"
      select="document('customXml/item1.xml')/b:Sources/b:Source[b:Tag = $TextIdentifier]"/>

    <xsl:variable name="BibliographyType" select="$Path/b:SourceType"/>

    <xsl:variable name="LastName">
      <xsl:choose>
        <xsl:when test="$Path/b:Author/b:Author/b:NameList/b:Person/b:Last">
          <xsl:value-of select="$Path/b:Author/b:Author/b:NameList/b:Person/b:Last"/>
        </xsl:when>
        <xsl:when test="$Path/b:Author/b:Author/b:Corporate">
          <xsl:value-of select="$Path/b:Author/b:Author/b:Corporate"/>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="FirstName">
      <xsl:choose>
        <xsl:when test="$Path/b:Author/b:Author/b:NameList/b:Person/b:First">
          <xsl:value-of select="$Path/b:Author/b:Author/b:NameList/b:Person/b:First"/>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="Middle">
      <xsl:choose>
        <xsl:when test="$Path/b:Author/b:Author/b:NameList/b:Person/b:Middle">
          <xsl:value-of select="$Path/b:Author/b:Author/b:NameList/b:Person/b:Middle"/>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="Author">
      <xsl:choose>
        <xsl:when test="$LastName != '' and $FirstName != '' and $Middle != ''">
          <xsl:value-of select="concat($LastName, ' ', $FirstName,' ', $Middle)"/>
        </xsl:when>
        <xsl:when test="$LastName != '' and $FirstName != ''">
          <xsl:value-of select="concat($LastName, ' ', $FirstName)"/>
        </xsl:when>
        <xsl:when test="$LastName != '' and $Middle != ''">
          <xsl:value-of select="concat($LastName, ' ', $Middle)"/>
        </xsl:when>
        <xsl:when test="$FirstName != '' and $Middle != ''">
          <xsl:value-of select="concat($FirstName,' ', $Middle)"/>
        </xsl:when>
        <xsl:when test="$LastName != ''">
          <xsl:value-of select="$LastName"/>
        </xsl:when>
        <xsl:when test="$FirstName != ''">
          <xsl:value-of select="$FirstName"/>
        </xsl:when>
        <xsl:when test="$Middle != ''">
          <xsl:value-of select="$Middle"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="City">
      <xsl:value-of select="$Path/b:City"/>
    </xsl:variable>

    <xsl:variable name="StateProvince">
      <xsl:value-of select="$Path/b:StateProvince"/>
    </xsl:variable>

    <xsl:variable name="CountryRegion">
      <xsl:value-of select="$Path/b:CountryRegion"/>
    </xsl:variable>

    <xsl:variable name="Address">
      <xsl:choose>
        <xsl:when test="$City != '' and $StateProvince != '' and $CountryRegion != ''">
          <xsl:value-of select="concat($City,' ',$StateProvince,' ',$CountryRegion)"/>
        </xsl:when>
        <xsl:when test="$City != '' and $StateProvince != ''">
          <xsl:value-of select="concat($City,' ',$StateProvince)"/>
        </xsl:when>
        <xsl:when test="$City != '' and $CountryRegion != ''">
          <xsl:value-of select="concat($City,' ',$CountryRegion)"/>
        </xsl:when>
        <xsl:when test="$StateProvince != '' and $CountryRegion != ''">
          <xsl:value-of select="concat($StateProvince,' ',$CountryRegion)"/>
        </xsl:when>
        <xsl:when test="$City != ''">
          <xsl:value-of select="$City"/>
        </xsl:when>
        <xsl:when test="$StateProvince != ''">
          <xsl:value-of select="$StateProvince"/>
        </xsl:when>
        <xsl:when test="$CountryRegion != ''">
          <xsl:value-of select="$CountryRegion"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    
    <xsl:variable name="TextIdent">
      <xsl:value-of select="$LastName"/>  
      <xsl:if test="$Path/b:Year">
      <xsl:text>, </xsl:text>
      <xsl:value-of select="$Path/b:Year"/>
      </xsl:if>
    </xsl:variable>

    <text:bibliography-mark>
      <xsl:attribute name="text:identifier">
        <xsl:value-of select="$TextIdent"/>
      </xsl:attribute>
      <xsl:attribute name="text:bibliography-type">
        <xsl:choose>
          <xsl:when test="$BibliographyType = 'Book'">
            <xsl:text>book</xsl:text>
          </xsl:when>
          <xsl:when test="$BibliographyType = 'JournalArticle'">
            <xsl:text>article</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>book</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:if test="$Author">
        <xsl:attribute name="text:author">
          <xsl:value-of select="$Author"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Path/b:Title">
        <xsl:attribute name="text:title">
          <xsl:value-of select="$Path/b:Title"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Path/b:Year">
        <xsl:attribute name="text:year">
          <xsl:value-of select="$Path/b:Year"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Path/b:Publisher">
        <xsl:attribute name="text:publisher">
          <xsl:value-of select="$Path/b:Publisher"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Address">
        <xsl:attribute name="text:address">
          <xsl:value-of select="$Address"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Path/b:Volume">
        <xsl:attribute name="text:volume">
          <xsl:value-of select="$Path/b:Volume"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Path/b:StandardNumber">
        <xsl:attribute name="text:number">
          <xsl:value-of select="$Path/b:StandardNumber"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Path/b:Pages">
        <xsl:attribute name="text:pages">
          <xsl:value-of select="$Path/b:Pages"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$Path/b:Edition">
        <xsl:attribute name="text:edition">
          <xsl:value-of select="$Path/b:Edition"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:value-of select="$LastName"/>
    </text:bibliography-mark>
  </xsl:template>

  <xsl:template name="InsertHyperlinkField">
    <xsl:param name="parentRunNode"/>
    <xsl:if test="$parentRunNode">
      <xsl:for-each select="$parentRunNode">
        <xsl:call-template name="InsertHyperlink"/>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertField">
    <xsl:choose>
      <!-- default scenario - catch beginning of field instruction. Other runs ignored (handled by first w:instrText processing). -->
      <xsl:when test="preceding::*[1][self::w:fldChar[@w:fldCharType='begin']] ">
        <xsl:apply-templates select="w:instrText[1]"/>
      </xsl:when>
      <xsl:otherwise>
        <!--  the same hyperlink can be in more then one paragraph so print as seperate text:a for each run which is in hyperlink field (between w:fldChar begin - end)-->
        <xsl:apply-templates select="preceding::w:instrText[1][contains(.,'HYPERLINK')]">
          <xsl:with-param name="parentRunNode" select="."/>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertEditingCycles">
    <text:editing-cycles>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:editing-cycles>
  </xsl:template>

  <!--editing cycles number REVNUM, DOCPROPERTY RevisionNumber, INFO RevNum-->
  <xsl:template
    match="w:fldSimple[contains(@w:instr,'REVNUM') or contains(@w:instr, 'RevisionNumber') or contains(@w:instr, 'RevNum')]"
    mode="fields">
    <xsl:call-template name="InsertEditingCycles"/>
  </xsl:template>

  <!--FILENAME, INFO FileName-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'FILENAME') or contains(@w:instr, 'FileName')]"
    mode="fields">
    <xsl:call-template name="InsertFileName"/>
  </xsl:template>

  <xsl:template name="InsertFileName">
    <text:file-name text:display="name">
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:file-name>
  </xsl:template>

  <!--KEYWORDS and DOCPROPERTY Keywords-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'KEYWORDS') or contains(@w:instr,'Keywords')]"
    mode="fields">
    <xsl:call-template name="InsertKeywords"/>
  </xsl:template>

  <xsl:template name="InsertKeywords">
    <text:keywords>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:keywords>
  </xsl:template>

  <!-- DOCPROPERTY Paragraphs, INFO Paragraphs-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'Company')]" mode="fields">
    <xsl:call-template name="InsertCompany"/>
  </xsl:template>

  <xsl:template name="InsertCompany">
    <text:sender-company>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:sender-company>
  </xsl:template>


  <xsl:template match="w:fldSimple[contains(@w:instr,'USERADDRESS')]" mode="fields">
    <xsl:call-template name="InsertUserAddress"/>
  </xsl:template>

  <xsl:template name="InsertUserAddress">
    <text:sender-street>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:sender-street>
  </xsl:template>

  <!--TEMPLATE, DOCPROPERTY Template-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'TEMPLATE')or contains(@w:instr, 'Template')]"
    mode="fields">
    <xsl:call-template name="InsertTemplate"/>
  </xsl:template>

  <xsl:template name="InsertTemplate">
    <text:template-name text:display="name">
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:template-name>
  </xsl:template>

  <!--NUMWORDS, DOCPROPERTY Words-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'NUMWORDS') or contains(@w:instr, 'Words')]"
    mode="fields">
    <xsl:call-template name="InsertWordCount"/>
  </xsl:template>

  <xsl:template name="InsertWordCount">
    <text:word-count>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:word-count>
  </xsl:template>

  <!--NUMCHARS and DOCPROPERTY Characters, INFO NumChars-->
  <xsl:template
    match="w:fldSimple[contains(@w:instr,'NUMCHARS') 
    or contains(@w:instr, 'Characters') or contains(@w:instr, 'NumChars')]"
    mode="fields">
    <xsl:call-template name="InsertCharacterCount"/>
  </xsl:template>

  <xsl:template name="InsertCharacterCount">
    <text:character-count>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:character-count>> </xsl:template>

  <!-- DOCPROPERTY Paragraphs, INFO Paragraphs-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'Paragraphs')]" mode="fields">
    <xsl:call-template name="InsertParagraphCount"/>
  </xsl:template>

  <xsl:template name="InsertParagraphCount">
    <text:paragraph-count>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:paragraph-count>
  </xsl:template>

  <!-- COMMENTS, DOCPROPERTY Comments-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'COMMENTS') or contains(@w:instr, 'Comments')]"
    mode="fields">
    <xsl:call-template name="InsertComments"/>
  </xsl:template>

  <xsl:template name="InsertComments">
    <text:description>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:description>
  </xsl:template>
</xsl:stylesheet>

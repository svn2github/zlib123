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
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" exclude-result-prefixes="w">

  <!-- Date and Time Fields -->
  <xsl:template name="InsertDateField">
    <xsl:param name="dateText"/>
    <text:date>
      <xsl:attribute name="style:data-style-name">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
      <xsl:attribute name="text:date-value">
        <xsl:value-of select="$dateText"/>
      </xsl:attribute>
      <xsl:value-of select="$dateText"/>
    </text:date>
  </xsl:template>

  <xsl:template name="InsertTimeField">
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
        <xsl:when test="$fieldType = 'XE' ">
          <xsl:call-template name="InsertIndexMark">
            <xsl:with-param name="instrText" select="following-sibling::instrText[1]"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'SET' ">
          <xsl:call-template name="InsertUserVariable">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'HYPERLINK' ">
          <xsl:call-template name="InsertHyperlinkField">
            <xsl:with-param name="parentRunNode" select="$parentRunNode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'REF' ">
          <xsl:call-template name="InsertCrossReference">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'DATE'">
          <xsl:call-template name="InsertDateField">
            <xsl:with-param name="dateText" select="$fieldCode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'NUMPAGE'">
          <xsl:call-template name="InsertPageCount"/>
        </xsl:when>
        <xsl:when test="$fieldType = 'PAGE' ">
          <xsl:call-template name="InsertPageNumberField"/>
        </xsl:when>
        <xsl:when test="$fieldType = 'TIME' ">
          <xsl:call-template name="InsertTimeField">
            <xsl:with-param name="timeText" select="$fieldCode"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$fieldType = 'USERINITIALS' ">
          <xsl:call-template name="InsertUserInitials"/>
        </xsl:when>
        <xsl:when test="$fieldType = 'USERNAME' ">
          <xsl:call-template name="InsertUserName"/>
        </xsl:when>
        <xsl:when test="$fieldType = 'AUTHOR'">
          <xsl:call-template name="InsertAuthor"/>
        </xsl:when>
        <!--caption field  from which Index of Figures is created -->
        <xsl:when test="$fieldType = 'SEQ' ">
          <xsl:call-template name="InsertSequence">
            <xsl:with-param name="fieldCode" select="$fieldCode"/>
            <xsl:with-param name="sequenceContext" select="following::w:r[w:t][1]"/>
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
      <xsl:when test="$fieldType = 'DATE' ">
        <xsl:call-template name="InsertDateStyle">
          <xsl:with-param name="dateText" select="$fieldCode"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$fieldType = 'TIME' ">
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
        test="$instrText/parent::w:r/following-sibling::*[self::w:r[w:instrText] or self::w:fldChar[@w:fldCharType = 'end']][1][self::w:r[w:instrText]]">
        <xsl:call-template name="BuildFieldCode">
          <!-- we know now that first run having instrText is before end of field -->
          <xsl:with-param name="instrText"
            select="$instrText/parent::w:r/following-sibling::*[self::w:r and w:instrText][1]/w:instrText[1]"/>
          <xsl:with-param name="fieldCode" select="concat($fieldCode, $instrText/text())"/>
        </xsl:call-template>
      </xsl:when>
      <!-- if next paragraph with instrText, before end of field :
      Find first paragraph having instrText before end of field, and then first run having instrText before end of field -->
      <xsl:when
        test="$instrText/ancestor::w:p/following-sibling::*[self::w:p[w:r/w:instrText or w:fldChar/@w:fldCharType = 'end']][1]/*[self::w:r[w:instrText] or self::w:fldChar[@w:fldCharType = 'end']][1][self::w:r[w:instrText]]">
        <xsl:call-template name="BuildFieldCode">
          <!-- we know now that first run having instrText is before end of field -->
          <xsl:with-param name="instrText"
            select="$instrText/ancestor::w:p/following-sibling::*[self::w:p[w:r/w:instrText]][1]/w:r/w:instrText"/>
          <xsl:with-param name="fieldCode" select="concat($fieldCode, $instrText/text())"/>
        </xsl:call-template>
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
    <text:alphabetical-index-mark>
      <xsl:attribute name="text:string-value">
        <xsl:value-of select="$instrText/text()"/>
      </xsl:attribute>
    </text:alphabetical-index-mark>
  </xsl:template>

  <xsl:template name="InsertCrossReference">
    <xsl:param name="fieldCode"/>
    <text:bookmark-ref text:reference-format="text">
      <xsl:attribute name="text:ref-name">
        <xsl:value-of select="substring-before(substring-after($fieldCode,'REF '),' \')"/>
      </xsl:attribute>
      <xsl:apply-templates select="following::w:t[1]/ancestor::w:r/child::node()"/>
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

  <!--document title-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'TITLE')]" mode="fields">
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

  <xsl:template match="w:fldSimple[contains(@w:instr,'AUTHOR')]" mode="fields">
    <xsl:call-template name="InsertAuthor"/>
  </xsl:template>

  <xsl:template name="InsertAuthor">
    <text:initial-creator>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:initial-creator>
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
            select="substring-before(substring-after(./@w:instr,'Heading '),'&quot;')"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:chapter>
  </xsl:template>

  <!--document subject-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'SUBJECT')]" mode="fields">
    <text:subject>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:subject>
  </xsl:template>

  <xsl:template match="w:fldSimple">
    <text:span text:style-name="{generate-id(w:r)}">
      <xsl:apply-templates select="." mode="fields"/>
    </text:span>
  </xsl:template>

  <xsl:template match="w:fldSimple[contains(@w:instr,'DATE')]" mode="fields">
    <xsl:call-template name="InsertDateField">
      <xsl:with-param name="dateText" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:fldSimple[contains(@w:instr,'TIME')]" mode="fields">
    <xsl:call-template name="InsertTimeField">
      <xsl:with-param name="timeText" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <!--page number-->
  <xsl:template
    match="w:fldSimple[contains(@w:instr,'PAGE ') and not(contains(@w:instr,'NUMPAGE'))]"
    mode="fields">
    <xsl:call-template name="InsertPageNumberField">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <!--page-count-->
  <xsl:template match="w:fldSimple[contains(@w:instr,'NUMPAGE')]" mode="fields">
    <xsl:call-template name="InsertPageCount">
      <xsl:with-param name="fieldCode" select="@w:instr"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="InsertPageCount">
    <text:page-count>
      <xsl:apply-templates select="w:r/child::node()"/>
    </text:page-count>
  </xsl:template>


  <xsl:template match="w:fldSimple[contains(@w:instr,'DATE')]" mode="automaticstyles">
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

    <text:sequence text:ref-name="{concat('ref',concat($refType,number($sequenceContext/w:t)-1))}" text:name="{$refType}" text:formula="{concat(concat('ooow:',$refType),'+1')}">
      <xsl:apply-templates select="$sequenceContext/child::node()"/>
    </text:sequence>
  </xsl:template>

  <xsl:template name="InsertDateStyle">
    <xsl:param name="dateText"/>
    <xsl:variable name="FormatDate">
      <xsl:value-of
        select="substring-before(substring-after($dateText, '&quot;'), '&quot;')"/>
    </xsl:variable>
    <xsl:call-template name="InsertDate">
      <xsl:with-param name="FormatDate">
        <xsl:value-of select="$FormatDate"/>
      </xsl:with-param>
      <xsl:with-param name="ParamField">
        <xsl:text>DATE</xsl:text>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:fldSimple[contains(@w:instr,'TIME')]" mode="automaticstyles">
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
    <xsl:call-template name="InsertDate">
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

  <xsl:template name="InsertDate">
    <xsl:param name="FormatDate"/>
    <xsl:param name="ParamField"/>
    <number:date-style>
      <xsl:attribute name="style:name">
        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>
      <xsl:call-template name="InsertFormatDate">
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

  <xsl:template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
        <xsl:call-template name="InsertFormatDate">
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
            <xsl:call-template name="InsertFormatDate">
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
            <xsl:call-template name="InsertFormatDate">
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
  <xsl:template name="InsertPageNumberField">
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
  <xsl:template name="TextBibliographyMark">
    <xsl:param name="TextIdentifier"/>

    <xsl:variable name="Path"
      select="document('customXml/item1.xml')/b:Sources/b:Source[b:Tag = $TextIdentifier]"/>

    <xsl:variable name="BibliographyType" select="$Path/b:SourceType"/>

    <xsl:variable name="LastName">
      <xsl:choose>
        <xsl:when test="$Path/b:Author/b:Author/b:NameList/b:Person/b:Last">
          <xsl:value-of select="$Path/b:Author/b:Author/b:NameList/b:Person/b:Last"/>
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

    <text:bibliography-mark>
      <xsl:attribute name="text:identifier">
        <xsl:value-of select="substring-before(substring-after(descendant::w:t, '('), ')')"/>
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
      <xsl:value-of select="substring-before(substring-after(descendant::w:t, '('), ')')"/>
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
        <!--  the same hyperlink can be in more then one paragraph so print as hyperlink each run which is in hyperlink field (between w:fldChar begin - end)-->
        <xsl:apply-templates select="preceding::w:instrText[1][contains(.,'HYPERLINK')]">
          <xsl:with-param name="parentRunNode" select="."/>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

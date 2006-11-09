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
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" exclude-result-prefixes="w">

  
  <xsl:template name="wInstrText">
    <xsl:if test="contains(w:instrText,'DATE')">
      <text:span text:style-name="{generate-id(self::node())}">
        <xsl:call-template name="InsertDataField"/>
      </text:span>
    </xsl:if>
    <xsl:if test="contains(w:instrText,'TIME') ">
      <text:span text:style-name="{generate-id(self::node())}">
        <xsl:call-template name="InsertTimeField"/>
      </text:span>
    </xsl:if>
  </xsl:template>
  
  <!-- Date and Time Fields -->
  
  <xsl:template name="InsertDataField">
    <text:date>
      <xsl:attribute name="style:data-style-name">
        <xsl:choose>
          <xsl:when test="parent::w:fldSimple">
            <xsl:value-of select="generate-id(parent::w:fldSimple)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="generate-id(w:instrText)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="text:date-value">
        <xsl:value-of select="."/>
      </xsl:attribute>
    </text:date>

  </xsl:template>

  <xsl:template name="InsertTimeField">

    <text:time>
      <xsl:attribute name="style:data-style-name">
        <xsl:choose>
          <xsl:when test="parent::w:fldSimple">
            <xsl:value-of select="generate-id(parent::w:fldSimple)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="generate-id(w:instrText)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="text:time-value">
        <xsl:value-of select="."/>
      </xsl:attribute>
    </text:time>

  </xsl:template>
  
  <!-- ignore text inside a field code -->
  <xsl:template match="w:instrText"/>

  <xsl:template match="w:instrText" mode="automaticstyles">

    <xsl:choose>
      <xsl:when test="contains(., 'DATE')">
        <xsl:variable name="FormatDate">
          <xsl:value-of select="substring-before(substring-after(., '&quot;'), '&quot;')"/>
        </xsl:variable>
        <xsl:call-template name="InsertDate">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="$FormatDate"/>
          </xsl:with-param>
          <xsl:with-param name="ParamField">
            <xsl:text>DATE</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="contains(., 'TIME')">
        <xsl:variable name="FormatDate">
          <xsl:value-of select="substring-before(substring-after(., '&quot;'), '&quot;')"/>
        </xsl:variable>
        <xsl:call-template name="InsertDate">
          <xsl:with-param name="FormatDate">
            <xsl:value-of select="$FormatDate"/>
          </xsl:with-param>
          <xsl:with-param name="ParamField">
            <xsl:text>TIME</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <!--  Insert Fields  -->
  
  <xsl:template name="wfldSimple">
    <xsl:choose>
      <xsl:when test="contains(parent::w:fldSimple/@w:instr, 'DATE')">
        <text:span text:style-name="{generate-id(self::node())}">
          <xsl:call-template name="InsertDataField"/>
        </text:span>
      </xsl:when>
      <xsl:when test="contains(parent::w:fldSimple/@w:instr, 'TIME')">
        <text:span text:style-name="{generate-id(self::node())}">
          <xsl:call-template name="InsertTimeField"/>
        </text:span>
      </xsl:when>
      <xsl:when
        test="contains(parent::w:fldSimple/@w:instr, 'NUMPAGES') or contains(parent::w:fldSimple/@w:instr, 'PAGE')">
        <xsl:call-template name="PageNumberField"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>


  <xsl:template match="w:fldSimple" mode="automaticstyles">
    
    <xsl:if test="contains(@w:instr,'DATE')">
      <xsl:variable name="FormatDate">
        <xsl:value-of
          select="substring-before(substring-after(@w:instr, '&quot;'), '&quot;')"/>
      </xsl:variable>
      <xsl:call-template name="InsertDate">
        <xsl:with-param name="FormatDate">
          <xsl:value-of select="$FormatDate"/>
        </xsl:with-param>
        <xsl:with-param name="ParamField">
          <xsl:text>DATE</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>

    <xsl:if test="contains(@w:instr,'TIME')">
      <xsl:variable name="FormatDate">
        <xsl:value-of
          select="substring-before(substring-after(@w:instr, '&quot;'), '&quot;')"/>
      </xsl:variable>
      <xsl:call-template name="InsertDate">
        <xsl:with-param name="FormatDate">
          <xsl:value-of select="$FormatDate"/>
        </xsl:with-param>
        <xsl:with-param name="ParamField">
          <xsl:text>TIME</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
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

  <!-- Page Number Field -->

  <xsl:template match="w:sdt/w:sdtContent/w:fldSimple">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template name="PageNumberField">
    <xsl:variable name="WInstr">
      <xsl:value-of select="parent::w:fldSimple/@w:instr"/>
    </xsl:variable>
    <text:page-number>
      <xsl:attribute name="style:num-format">
        <xsl:choose>
          <xsl:when test="contains($WInstr, 'Arabic')">Arabic</xsl:when>
          <xsl:when test="contains($WInstr, 'alphabetic')">a</xsl:when>
          <xsl:when test="contains($WInstr, 'ALPHABETIC')">A</xsl:when>
          <xsl:when test="contains($WInstr, 'roman')">i</xsl:when>
          <xsl:when test="contains($WInstr, 'ROMAN')">I</xsl:when>
          <xsl:otherwise>Arabic</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>      
    </text:page-number>
  </xsl:template>




</xsl:stylesheet>

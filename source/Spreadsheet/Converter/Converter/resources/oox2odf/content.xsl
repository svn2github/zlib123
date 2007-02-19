<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="measures.xsl"/>
  <xsl:import href="styles.xsl"/>

  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls>
        <xsl:call-template name="InsertFonts"/>
      </office:font-face-decls>
      <office:automatic-styles>
        <xsl:call-template name="InsertColumnStyles"/>
        <xsl:call-template name="InsertRowStyles"/>
        <xsl:call-template name="InsertCellStyles"/>
        <xsl:call-template name="InsertStyleTableProperties"/>
      </office:automatic-styles>
      <xsl:call-template name="InsertSheets"/>
    </office:document-content>
  </xsl:template>

  <xsl:template name="InsertSheets">
    <office:body>
      <office:spreadsheet>
        <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
          <table:table>

            <!-- Insert Table (Sheet) Name -->

            <xsl:attribute name="table:name">
              <xsl:value-of select="@name"/>
            </xsl:attribute>

            <!-- Insert Table Style Name (style:table-properties) -->

            <xsl:attribute name="table:style-name">
              <xsl:value-of select="generate-id()"/>
            </xsl:attribute>

            <xsl:call-template name="InsertSheetContent">
              <xsl:with-param name="sheet">
                <xsl:call-template name="GetTarget">
                  <xsl:with-param name="id">
                    <xsl:value-of select="@r:id"/>
                  </xsl:with-param>
                  <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </table:table>
        </xsl:for-each>
      </office:spreadsheet>
    </office:body>
  </xsl:template>

  <xsl:template name="InsertSheetContent">
    <xsl:param name="sheet"/>

    <xsl:call-template name="InsertColumns">
      <xsl:with-param name="sheet" select="$sheet"/>
    </xsl:call-template>

    <xsl:apply-templates select="document(concat('xl/',$sheet))/e:worksheet"/>
    <xsl:if test="not(document(concat('xl/',$sheet))/e:worksheet/e:sheetData/e:row/e:c/e:v)">
      <table:table-row>
        <table:table-cell/>
      </table:table-row>
    </xsl:if>
  </xsl:template>

  <xsl:template match="e:row">

    <xsl:variable name="lastCellColumnNumber">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="e:c[last()]/@r"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <!-- if first rows are empty-->
      <xsl:when test="position()=1">
        <xsl:if test="@r>1">
          <table:table-row table:number-rows-repeated="{@r - 1}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <!-- if there's a gap between rows -->
        <xsl:if test="preceding::e:row[1]/@r &lt;  @r - 1">
          <table:table-row>
            <xsl:attribute name="table:number-rows-repeated">
              <xsl:value-of select="@r -1 - preceding::e:row[1]/@r"/>
            </xsl:attribute>
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

    <table:table-row>
      <xsl:apply-templates/>
      <xsl:if test="$lastCellColumnNumber &lt; 256">
        <table:table-cell table:number-columns-repeated="{256 - $lastCellColumnNumber}"/>
      </xsl:if>
    </table:table-row>
  </xsl:template>

  <xsl:template match="e:c">

    <xsl:variable name="this" select="."/>

    <xsl:variable name="colNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="prevCellColNum">
      <xsl:choose>
        <xsl:when
          test="preceding::e:c[1] and generate-id(preceding::e:c[1]/parent::node())=generate-id(parent::node())">
          <xsl:call-template name="GetColNum">
            <xsl:with-param name="cell">
              <xsl:value-of select="preceding::e:c[1]/@r"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>-1</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- insert blank cells before this one-->
    <xsl:choose>
      <xsl:when test="position()=1 and $colNum>1">
        <table:table-cell>
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$colNum - 1"/>
          </xsl:attribute>
        </table:table-cell>
      </xsl:when>
      <xsl:when test="position()>1 and $colNum>$prevCellColNum+1">
        <table:table-cell>
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$colNum - $prevCellColNum - 1"/>
          </xsl:attribute>
        </table:table-cell>
      </xsl:when>
    </xsl:choose>

    <!-- insert this one cell-->
    <table:table-cell>
      <xsl:choose>
        <xsl:when test="@t='s'">
          <xsl:attribute name="office:value-type">
            <xsl:text>string</xsl:text>
          </xsl:attribute>
          <xsl:if test="@s">
            <xsl:attribute name="table:style-name">
              <xsl:value-of
                select="generate-id(document('xl/styles.xml')/e:styleSheet/e:cellXfs/e:xf[position() = $this/@s + 1])"
              />
            </xsl:attribute>
          </xsl:if>
          <xsl:variable name="id">
            <xsl:value-of select="e:v"/>
          </xsl:variable>
          <text:p>
            <xsl:value-of select="document('xl/sharedStrings.xml')/e:sst/e:si[position()=$id+1]/e:t"
            />
          </text:p>
        </xsl:when>
        <xsl:otherwise>
          <text:p>
            <xsl:value-of select="e:v"/>
          </text:p>
        </xsl:otherwise>
      </xsl:choose>
    </table:table-cell>
  </xsl:template>

  <!-- gets a column number from cell coordinates -->
  <xsl:template name="GetColNum">
    <xsl:param name="cell"/>
    <xsl:param name="columnId"/>

    <xsl:choose>
      <!-- when whole literal column id has been extracted than convert alphabetic index to number -->
      <xsl:when test="number(substring($cell,1,1))">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="$columnId"/>
        </xsl:call-template>
      </xsl:when>
      <!--  recursively extract literal column id (i.e if $cell='GH15' it will return 'GH') -->
      <xsl:otherwise>
        <xsl:call-template name="GetColNum">
          <xsl:with-param name="cell" select="substring-after($cell,substring($cell,1,1))"/>
          <xsl:with-param name="columnId" select="concat($columnId,substring($cell,1,1))"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- translates literal index to number -->
  <xsl:template name="GetAlphabeticPosition">
    <xsl:param name="literal"/>
    <xsl:param name="number" select="0"/>
    <xsl:param name="level" select="0"/>

    <xsl:variable name="lastCharacter">
      <xsl:value-of select="substring($literal,string-length($literal))"/>
    </xsl:variable>

    <xsl:variable name="lastCharacterPosition">
      <xsl:call-template name="CharacterToPosition">
        <xsl:with-param name="character" select="$lastCharacter"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="power">
      <xsl:call-template name="Power">
        <xsl:with-param name="base" select="26"/>
        <xsl:with-param name="exponent" select="$level"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="string-length($literal)>1">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="substring-before($literal,$lastCharacter)"/>
          <xsl:with-param name="level" select="$level+1"/>
          <xsl:with-param name="number">
            <xsl:value-of select="$lastCharacterPosition*$power + $number"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$lastCharacterPosition*$power + $number"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- returns position in alphabet of a single character-->
  <xsl:template name="CharacterToPosition">
    <xsl:param name="character"/>

    <xsl:choose>
      <xsl:when test="$character='A'">1</xsl:when>
      <xsl:when test="$character='B'">2</xsl:when>
      <xsl:when test="$character='C'">3</xsl:when>
      <xsl:when test="$character='D'">4</xsl:when>
      <xsl:when test="$character='E'">5</xsl:when>
      <xsl:when test="$character='F'">6</xsl:when>
      <xsl:when test="$character='G'">7</xsl:when>
      <xsl:when test="$character='H'">8</xsl:when>
      <xsl:when test="$character='I'">9</xsl:when>
      <xsl:when test="$character='J'">10</xsl:when>
      <xsl:when test="$character='K'">11</xsl:when>
      <xsl:when test="$character='L'">12</xsl:when>
      <xsl:when test="$character='M'">13</xsl:when>
      <xsl:when test="$character='N'">14</xsl:when>
      <xsl:when test="$character='O'">15</xsl:when>
      <xsl:when test="$character='P'">16</xsl:when>
      <xsl:when test="$character='Q'">17</xsl:when>
      <xsl:when test="$character='R'">18</xsl:when>
      <xsl:when test="$character='S'">19</xsl:when>
      <xsl:when test="$character='T'">20</xsl:when>
      <xsl:when test="$character='U'">21</xsl:when>
      <xsl:when test="$character='V'">22</xsl:when>
      <xsl:when test="$character='W'">23</xsl:when>
      <xsl:when test="$character='X'">24</xsl:when>
      <xsl:when test="$character='Y'">25</xsl:when>
      <xsl:when test="$character='Z'">26</xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- calculates power function -->
  <xsl:template name="Power">
    <xsl:param name="base"/>
    <xsl:param name="exponent"/>
    <xsl:param name="value1" select="$base"/>

    <xsl:choose>
      <xsl:when test="$exponent = 0">
        <xsl:text>1</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$exponent &gt; 1">
            <xsl:call-template name="Power">
              <xsl:with-param name="base">
                <xsl:value-of select="$base"/>
              </xsl:with-param>
              <xsl:with-param name="exponent">
                <xsl:value-of select="$exponent -1"/>
              </xsl:with-param>
              <xsl:with-param name="value1">
                <xsl:value-of select="$value1 * $base"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$value1"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertColumns">
    <xsl:param name="sheet"/>

    <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet/e:cols/e:col">
      <!-- insert blank columns before this one-->
      <xsl:choose>
        <xsl:when test="position()=1 and @min>1">
          <table:table-column
            table:style-name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
            table:number-columns-repeated="{@min - 1}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of
                select="generate-id(document('xl/styles.xml')/e:styleSheet/e:cellXfs/e:xf[1])"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:when>
        <xsl:when test="preceding::e:col[1]/@max &lt; @min - 1">
          <table:table-column
            table:style-name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
            table:number-columns-repeated="{@min - preceding::e:col[1]/@max - 1}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of
                select="generate-id(document('xl/styles.xml')/e:styleSheet/e:cellXfs/e:xf[1])"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:when>
      </xsl:choose>

      <!-- insert this one column -->
      <table:table-column table:style-name="{generate-id(.)}">
        <xsl:if test="not(@min = @max)">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="@max - @min"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name="table:default-cell-style-name">
          <xsl:value-of
            select="generate-id(document('xl/styles.xml')/e:styleSheet/e:cellXfs/e:xf[1])"/>
        </xsl:attribute>
      </table:table-column>
    </xsl:for-each>

    <!-- apply default column style for last columns which style wasn't changed -->
    <xsl:choose>
      <xsl:when test="not(document(concat('xl/',$sheet))/e:worksheet/e:cols/e:col)">
        <table:table-column
          table:style-name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
          table:number-columns-repeated="256">
          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of
              select="generate-id(document('xl/styles.xml')/e:styleSheet/e:cellXfs/e:xf[1])"/>
          </xsl:attribute>
        </table:table-column>
      </xsl:when>
      <xsl:otherwise>
        <table:table-column
          table:style-name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
          table:number-columns-repeated="{256 - document(concat('xl/',$sheet))/e:worksheet/e:cols/e:col[last()]/@max}">
          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of
              select="generate-id(document('xl/styles.xml')/e:styleSheet/e:cellXfs/e:xf[1])"/>
          </xsl:attribute>
        </table:table-column>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="ConvertFromCharacters">
    <xsl:param name="value"/>

    <xsl:variable name="defaultFontSize">
      <xsl:value-of select="document('xl/styles.xml')/e:styleSheet/e:fonts/e:font[1]/e:sz/@val"/>
    </xsl:variable>

    <!-- formula below is true only for proportional fonts -->
    <xsl:variable name="avgDigitWidth">
      <xsl:value-of select="floor($defaultFontSize * 0.66666)"/>
    </xsl:variable>

    <xsl:call-template name="ConvertToCentimeters">
      <xsl:with-param name="length">
        <xsl:value-of select="concat(round(($avgDigitWidth * $value) + 5),'px')"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="InsertFonts">
    <xsl:for-each select="document('xl/styles.xml')/e:styleSheet/e:fonts/e:font">
      <style:font-face style:name="{e:name/@val}" svg:font-family="{e:name/@val}"> </style:font-face>

    </xsl:for-each>
  </xsl:template>

</xsl:stylesheet>

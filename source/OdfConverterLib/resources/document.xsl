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
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
	exclude-result-prefixes="office text table fo style">

	<xsl:strip-space elements='*'/>
	<xsl:preserve-space elements='text:p'/>
	<xsl:preserve-space elements="text:span"/>

	<xsl:key name="style" match="office:automatic-styles/style:style|office:automatic-styles/style:style" use="@style:name"/>
	
	<xsl:variable name="type">dxa</xsl:variable>
	
	<xsl:template name="document">
		<w:document>
			<xsl:apply-templates select="document('content.xml')/office:document-content"/>
		</w:document>
	</xsl:template>
	
	<xsl:template name="subtable">
		<xsl:param name="node"/>
		<xsl:for-each select="$node/table:table-cell">
			<xsl:call-template name="table-cell"/>
		</xsl:for-each>
	</xsl:template>
	
	<xsl:template name="merged-rows">
		<xsl:param name="i" select="0"/>
		<xsl:param name="iterator"/>
		<xsl:variable name="test">
			<xsl:if test="$i > 0">
				<xsl:text>true</xsl:text>
			</xsl:if>
		</xsl:variable>
		<xsl:if test="$test='true'">
			<w:tr>
				<xsl:for-each select="table:table-cell">
					<xsl:choose>
						<xsl:when test="table:table[@table:is-sub-table='true']">
							<!-- table to process -->
							<xsl:call-template name="subtable">
								<xsl:with-param name="node" select="table:table/child::table:table-row[$iterator]"></xsl:with-param>
							</xsl:call-template>
						</xsl:when>
						<xsl:when test="@table:number-columns-spanned">
							<xsl:choose>
								<xsl:when test="$iterator = 1">
									<xsl:call-template name="table-cell">
										<xsl:with-param name="grid" select="round(number(@table:number-columns-spanned))"/>
										<xsl:with-param name="merge" select="1"/>									
									</xsl:call-template>
								</xsl:when>
								<xsl:otherwise>
									<xsl:call-template name="table-cell">
										<xsl:with-param name="grid" select="round(number(@table:number-columns-spanned))"/>
										<xsl:with-param name="merge" select="2"/>									
									</xsl:call-template>
								</xsl:otherwise>
							</xsl:choose>
							
						</xsl:when>
						<xsl:otherwise>
							<xsl:apply-templates select="."/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</w:tr>
			<xsl:call-template name="merged-rows">
				<xsl:with-param name="i" select="$i  -1"/>
				<xsl:with-param name="iterator" select="$iterator +1"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	
	<!--xsl:template match="text()" mode="document">
	</xsl:template-->
	
	<xsl:template match="office:body">
		<w:body>
			<xsl:apply-templates/>
		</w:body>
	</xsl:template>

	<xsl:template match="office:text">
		<xsl:apply-templates/>
	</xsl:template>

	<!-- headings -->
	
	<xsl:template match="text:h">
		<w:p>
			<w:pPr>
				<w:pStyle w:val="{@text:style-name}"/>
				<xsl:choose>
					<xsl:when test="not(@text:outline-level)">
						<w:outlineLvl w:val="0"/>
					</xsl:when>
					<xsl:when test="@text:outline-level &lt;= 9">
						<w:outlineLvl w:val="{@text:outline-level}"/>
					</xsl:when>
					<xsl:otherwise>
						<w:outlineLvl w:val="9"/>
					</xsl:otherwise>
				</xsl:choose>
			</w:pPr>
			<xsl:apply-templates mode="paragraph"/>
		</w:p>
	</xsl:template>

	<!-- paragraphs -->
	
	<xsl:template match="text:p">
		<xsl:param name="level" select="0"/>
    <xsl:message terminate="no">text:p</xsl:message>
		<xsl:choose>
			<xsl:when test="$level = 0">
				<w:p>
					<w:pPr>
						<w:pStyle w:val="{@text:style-name}"/>
						<xsl:variable name="style" select="@text:style-name"/>
						<xsl:for-each select="/office:document-content/office:automatic-styles/style:style">
							<xsl:if test="@style:name = $style">
								<xsl:if test="style:paragraph-properties/@fo:line-height ">
									<xsl:variable name="space" select="round(number(substring-before(style:paragraph-properties/@fo:line-height, '%')) * 24 div 10)"/>
									<w:spacing w:line="{$space}" w:lineRule="auto"/>
								</xsl:if>
							</xsl:if>
						</xsl:for-each>
					</w:pPr>
					<xsl:apply-templates mode="paragraph"/>
				</w:p>		
			</xsl:when>
			<xsl:otherwise>
				<!-- We are in a list -->
				<w:p>
					<w:pPr>
						<w:pStyle w:val='{@text:style-name}'/>
						<w:numPr>
							<w:ilvl w:val="{$level}"/>
							<xsl:variable name="id" select="ancestor::text:list/@text:style-name"/>
							<w:numId w:val="{count(key('list-style',$id)/preceding-sibling::text:list-style)+1}"/>
						</w:numPr>
					</w:pPr>
					<xsl:apply-templates mode="paragraph"/>
				</w:p>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- links -->
	
	<!--xsl:template match="text:a" mode="paragraph">
		<w:hyperlink r:id='{generate-id()}'>
			<w:r>
				<w:rPr>
					<w:rStyle w:val="{@text:style-name}"/>
				</w:rPr>
				<w:t>
					<xsl:apply-templates mode="hyperlink"/>
				</w:t>
			</w:r>
		</w:hyperlink>
	</xsl:template-->
	
	<xsl:template match="text:a" mode="paragraph">
		<w:hyperlink r:id='{generate-id()}'>
			<xsl:apply-templates mode="paragraph"/>
		</w:hyperlink>
	</xsl:template>
	
	<!-- TODO : find the best way to avoid code duplication -->
	<xsl:template match="text:tab-stop" mode="hyperlink">
		<w:tab/>
	</xsl:template>

	<xsl:template match="text:line-break" mode="hyperlink">
		<w:br/>
	</xsl:template>

	<xsl:template match="text()" mode="hyperlink">
		<xsl:value-of select="."/>
	</xsl:template>
	
	<xsl:template match="text:s" mode="hyperlink">
		<xsl:call-template name="extra-spaces"/>
	</xsl:template>
	
	<xsl:template match="text:span" mode="hyperlink">
		<w:rPr>
			<w:rStyle w:val="{@text:style-name}"/>
		</w:rPr>
		<w:t>
			<xsl:attribute name="xml:space">preserve</xsl:attribute>
			<xsl:apply-templates mode="hyperlink"/>
		</w:t>
	</xsl:template>
	
	<!-- lists -->
	
	<xsl:template match="text:list">
		<xsl:param name="level" select="-1"/>
       	<xsl:apply-templates select="text:list-item">
       		<xsl:with-param name="level" select="$level+1"/>
		</xsl:apply-templates>
	</xsl:template>
	
	<xsl:template match="text:list-item">
		<xsl:param name="level"/>
	  <xsl:choose>
  	  <xsl:when test="*[1][self::text:p]">
    		<w:p>
    			<w:pPr>
    				<w:pStyle w:val="{*[1][self::text:p]/@text:style-name}"/>
    				<w:numPr>
    					<w:ilvl w:val="{$level}"/>
    					<xsl:variable name="id" select="ancestor::text:list/@text:style-name"/>
    					<w:numId w:val="{count(key('list-style',$id)/preceding-sibling::text:list-style)+1}"/>
    				</w:numPr>
    			</w:pPr>
    			<!-- first paragraph -->
    			<xsl:apply-templates select='*[1][self::text:p]' mode="paragraph"/>
    		</w:p>
    		<!-- others (text:p or text:list) -->
    		<xsl:apply-templates select='*[position() != 1]'>
    			<xsl:with-param name="level" select="$level"/>
    		</xsl:apply-templates>
  	  </xsl:when>
	    <xsl:otherwise>
    	  <xsl:apply-templates>
    	    <xsl:with-param name="level" select="$level"/>
    	  </xsl:apply-templates>
	    </xsl:otherwise>
	  </xsl:choose>
	</xsl:template>

	<!-- tables -->
	
	<xsl:template match="table:table">
		<w:tbl>
			<w:tblPr>
				<w:tblStyle w:val="{@table:style-name}"></w:tblStyle>
				<w:tblW  w:type="{$type}">
					<xsl:attribute name="w:w">
						<xsl:call-template name="twips-measure">
							<xsl:with-param name="length" select="key('style', @table:style-name)/style:table-properties/@style:width"/>
						</xsl:call-template>
					</xsl:attribute>
				</w:tblW>
			</w:tblPr>
			<w:tblGrid>
				<xsl:apply-templates select="table:table-column"/>
			</w:tblGrid>
			<xsl:apply-templates select="table:table-rows|table:table-header-rows|table:table-row|table:table-header-row"/>
		</w:tbl>			
	</xsl:template>
  
	<xsl:template match="table:table-column">
		<xsl:param name="repeat" select="1"/>
		<!-- relative width not supported yet -->
		<w:gridCol>
			<xsl:attribute name="w:w">
				<xsl:call-template name="twips-measure">
					<xsl:with-param name="length" select="key('style', @table:style-name)/style:table-column-properties/@style:column-width"/>
				</xsl:call-template>
			</xsl:attribute>
		</w:gridCol>
		<xsl:if test="@table:number-columns-repeated ">
			<xsl:if test="@table:number-columns-repeated &gt; $repeat">
				<xsl:apply-templates select=".">
					<xsl:with-param name="repeat" select="$repeat + 1"/>
				</xsl:apply-templates>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	
	<xsl:template match="table:table-rows|table:table-header-rows">
		<xsl:apply-templates select="table:table-row|table:table-header-row"/>
	</xsl:template>
	
	<xsl:template match="table:table-row|table:table-header-row">
		<xsl:choose>
			<xsl:when test="table:table-cell/table:table/@table:is-sub-table='true'">
				<!-- merged cells -->
				<xsl:variable name="total_rows" select="count(table:table-cell/table:table[@table:is-sub-table='true']/table:table-row)"/>
				<xsl:variable name="subtables" select="count(table:table-cell/table:table[@table:is-sub-table='true'])"/>
				<xsl:call-template name="merged-rows">
					<xsl:with-param name="i" select="$total_rows div $subtables"/>
					<xsl:with-param name="iterator" select="1"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<w:tr>
					<w:trPr>
						<xsl:if test="name(parent::*) = 'table:table-header-rows'">
							<w:tblHeader/>
						</xsl:if>
						<!-- row styles -->
					</w:trPr>
					<xsl:apply-templates select="table:table-cell"/>
				</w:tr>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="table:table-cell" name="table-cell">
		<xsl:param name="grid" select="0"/>
		<xsl:param name="merge" select="0"/>	
		<w:tc>
			<w:tcPr>
				<!-- point on the cell style properties --> 
				<xsl:variable name="cellProp" select="key('style', @table:style-name)/style:table-cell-properties"/>
	
				<!-- @TODO : width of the cell -->
				<xsl:choose>
					<xsl:when test="$merge = 1">
						<w:gridSpan w:val="{$grid}"/>
						<w:vmerge w:val="restart"/>
					</xsl:when>
					<xsl:when test="$merge = 2">
						<w:gridSpan w:val="{$grid}"/>
						<w:vmerge/>
					</xsl:when>
				</xsl:choose>
				<w:tcBorders>
					<xsl:choose>
						<xsl:when test="$cellProp[@fo:border and @fo:border!='none']">
							<xsl:variable name="border" select="$cellProp/@fo:border"/>
							<!-- fo:border = "0.002cm solid #000000" -->
							<xsl:variable name="border-color" select="substring-after($border, '#')"/>
							<xsl:variable name="border-size">
								<xsl:call-template name="eightspoint-measure">
									<xsl:with-param name="length" select="substring-before($border,' ')"/>
								</xsl:call-template>
							</xsl:variable>
							<w:top w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:left w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:bottom w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:right w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:if test="$cellProp[@fo:border-top and @fo:border-top != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-top"/>
								<w:top w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:top>
							</xsl:if>
							<xsl:if test="$cellProp[@fo:border-left and @fo:border-left != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-left"/>
								<w:left w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:left>
							</xsl:if>
							<xsl:if test="$cellProp[@fo:border-bottom and @fo:border-bottom != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-bottom"/>
								<w:bottom w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:bottom>
							</xsl:if>
							<xsl:if test="$cellProp[@fo:border-right and @fo:border-right != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-right"/>
								<w:right w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:right>
							</xsl:if>
						</xsl:otherwise>
					</xsl:choose>
				</w:tcBorders>

				<xsl:choose>
					<xsl:when test="$cellProp[@fo:background-color]">
						<xsl:variable name="fill" select="$cellProp/@fo:background-color"/>
						<w:shd w:val="clear" w:color="auto" w:fill="{substring-after($fill, '#')}" />
					</xsl:when>
				</xsl:choose>

				<w:tcMar>
					<xsl:choose>
						<xsl:when test="$cellProp[@fo:padding and @fo:padding != 'none']">
							<xsl:variable name="padding">
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length" select="$cellProp/@fo:padding"/>
								</xsl:call-template>
							</xsl:variable>
							<w:top w:w="{$padding}" w:type="{$type}"/>
							<w:left w:w="{$padding}" w:type="{$type}"/>
							<w:bottom w:w="{$padding}" w:type="{$type}"/>
							<w:right w:w="{$padding}" w:type="{$type}"/>
						</xsl:when>
						<xsl:otherwise>
							<w:top w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-top"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:top>
							<w:left w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-left"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:left>
							<w:bottom w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-bottom"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:bottom> 
							<w:right w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-right"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:right>	
						</xsl:otherwise>
					</xsl:choose>
				</w:tcMar>
				
				<xsl:if test="$cellProp[@fo:vertical-align and @fo:vertical-align!='']">
					<w:vAlign w:val="{$cellProp/@fo:vertical-align}"/>
				</xsl:if>	
			</w:tcPr>
			<xsl:apply-templates/>
			<w:p/> <!-- must precede a w:tc, otherwise it crashes. Xml schema validation does not check this. -->
		</w:tc>
	</xsl:template>
	
	<!-- text and spaces -->
	
	<xsl:template match="text()" mode="paragraph">
		<w:r>
			<xsl:apply-templates select="." mode="text"/>
		</w:r>
	</xsl:template>
	
	
	<xsl:template match="text()" mode="text">
		<w:t>
				<xsl:attribute name="xml:space">preserve</xsl:attribute>
				<xsl:value-of select="."/>
			
				<!-- extra-spaces inclusion -->
				<xsl:if test="following-sibling::text:s[1]">
					<xsl:call-template name="extra-spaces">
						<xsl:with-param name="spaces" select="following-sibling::text:s[1]/@text:c"></xsl:with-param>
					</xsl:call-template>
				</xsl:if>
		</w:t>
	</xsl:template>
	

	<xsl:template match="text:span" mode="paragraph">
		<w:r>
			<w:rPr>
				<w:rStyle w:val="{@text:style-name}"/>
				<xsl:variable name="style" select="@text:style-name"/>
				<xsl:for-each select="/office:document-content/office:automatic-styles/style:style">
					<xsl:if test="@style:name = $style">
						<xsl:if test="style:text-properties/@style:text-position">
							<xsl:variable name="pos" select="substring-before(style:text-properties/@style:text-position, ' ')"/>
							<xsl:choose>
								<xsl:when test="$pos = 'sub'">
									<w:vertAlign w:val="subscript"/>
								</xsl:when>
								<xsl:when test="$pos = 'super'">
									<w:vertAlign w:val="superscript"/>
								</xsl:when>
							</xsl:choose>
						</xsl:if>
					</xsl:if>
				</xsl:for-each>
			</w:rPr>
			<xsl:apply-templates mode="text"/>
		</w:r>
	</xsl:template>

	<xsl:template match="text:tab-stop" mode="paragraph">
		<w:r><w:tab/><w:t/></w:r>
	</xsl:template>

	
	<xsl:template match="text:tab" mode="paragraph">
		<w:r><w:tab/></w:r>
	</xsl:template>
	
	
	<xsl:template match="text:line-break" mode="paragraph">
		<w:r><w:br/><w:t/></w:r>
	</xsl:template>
	
	<xsl:template match="text:line-break" mode="text">
		<w:br/>
	</xsl:template>


	<!-- Extra spaces management, courtesy of J. David Eisenberg -->
	<xsl:variable name="spaces" xml:space="preserve">                                       </xsl:variable>
	
	<xsl:template name="extra-spaces">
		<xsl:param name="spaces"/>
		<xsl:choose>
			<xsl:when test="$spaces">
				<xsl:call-template name="insert-spaces">
					<xsl:with-param name="n" select="$spaces"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text> </xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template name="insert-spaces">
		<xsl:param name="n"/>
		<xsl:choose>
			<xsl:when test="$n &lt;= string-length($spaces)">
				<xsl:value-of select="substring($spaces, 1, $n)" xml:space="preserve"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$spaces"/>
				<xsl:call-template name="insert-spaces">
					<xsl:with-param name="n">
						<xsl:value-of select="$n - string-length($spaces)"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- ignored -->
	<xsl:template match="text()"/>

</xsl:stylesheet>

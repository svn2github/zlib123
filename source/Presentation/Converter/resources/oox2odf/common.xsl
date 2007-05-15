<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="style svg fo">
	<xsl:template name="getColorCode">
		<xsl:param name="color"/>
		<xsl:param name ="lumMod"/>
		<xsl:param name ="lumOff"/>
		<xsl:choose>
			<!--White, Background1-->
			<xsl:when test ="($color = 'bg1' or contains($color,'bg1')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#FFFFFF'" />
			</xsl:when>
			<xsl:when test ="($color = 'bg1' or contains($color,'bg1')) and $lumMod='95000' and $lumOff=''">
				<xsl:value-of select ="'#f2f2f2'" />
			</xsl:when>
			<xsl:when test ="($color = 'bg1' or contains($color,'bg1')) and $lumMod='85000' and $lumOff=''">
				<xsl:value-of select ="'#d9d9d9'" />
			</xsl:when>
			<xsl:when test ="($color = 'bg1' or contains($color,'bg1')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#bfbfbf'" />
			</xsl:when>
			<xsl:when test ="($color = 'bg1' or contains($color,'bg1')) and $lumMod='65000' and $lumOff=''">
				<xsl:value-of select ="'#a6a6a6'"/>
			</xsl:when>
			<xsl:when test ="($color = 'bg1' or contains($color,'bg1')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#7f7f7f'"/>
			</xsl:when>

			<!--Black, Text 1-->
			<xsl:when test ="($color = 'tx1' or contains($color,'tx1')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#000000'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx1' or contains($color,'tx1')) and $lumMod='50000' and $lumOff='50000'">
				<xsl:value-of select ="'#7f7f7f'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx1' or contains($color,'tx1')) and $lumMod='65000' and $lumOff='35000'">
				<xsl:value-of select ="'#595959'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx1' or contains($color,'tx1')) and $lumMod='75000' and $lumOff='25000'">
				<xsl:value-of select ="'#404040'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx1' or contains($color,'tx1')) and $lumMod='85000' and $lumOff='15000'">
				<xsl:value-of select ="'#262626'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx1' or contains($color,'tx1')) and $lumMod='95000' and $lumOff='5000'">
				<xsl:value-of select ="'#0d0d0d'"/>
			</xsl:when>

			<!--Tan, Background 2-->
			<xsl:when test ="($color = 'bg2' or contains($color,'bg2')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#EEECE1'"/>
			</xsl:when>
			<xsl:when test ="($color = 'bg2' or contains($color,'bg2')) and $lumMod='90000' and $lumOff=''">
				<xsl:value-of select ="'#DDD9C3'"/>
			</xsl:when>
			<xsl:when test ="($color = 'bg2' or contains($color,'bg2')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#C4BD97'" />
			</xsl:when>
			<xsl:when test ="($color = 'bg2' or contains($color,'bg2')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#948A54'"/>
			</xsl:when>
			<xsl:when test ="($color = 'bg2' or contains($color,'bg2')) and $lumMod='25000' and $lumOff=''">
				<xsl:value-of select ="'#4A452A'"/>
			</xsl:when>
			<xsl:when test ="($color = 'bg2' or contains($color,'bg2')) and $lumMod='10000' and $lumOff=''">
				<xsl:value-of select ="'#1E1C11'" />
			</xsl:when>

			<!--Dark Blue, Text 2-->
			<xsl:when test ="($color = 'tx2' or contains($color,'tx2')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#1F497D'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx2' or contains($color,'tx2')) and $lumMod='20000' and $lumOff='80000'">
				<xsl:value-of select ="'#C6D9F1'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx2' or contains($color,'tx2')) and $lumMod='40000' and $lumOff='40000'">
				<xsl:value-of select ="'#8EB4E3'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx2' or contains($color,'tx2')) and $lumMod='60000' and $lumOff='40000'">
				<xsl:value-of select ="'#558ED5'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx2' or contains($color,'tx2')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#17375E'"/>
			</xsl:when>
			<xsl:when test ="($color = 'tx2' or contains($color,'tx2')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#10253F'"/>
			</xsl:when>

			<!--Blue, Accent 1-->
			<xsl:when test ="($color = 'accent1' or contains($color,'accent1')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#4F81BD'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent1' or contains($color,'accent1')) and $lumMod='20000' and $lumOff='80000'">
				<xsl:value-of select ="'#DCE6F2'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent1' or contains($color,'accent1')) and $lumMod='40000' and $lumOff='60000'">
				<xsl:value-of select ="'#B9CDE5'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent1' or contains($color,'accent1')) and $lumMod='60000' and $lumOff='40000'">
				<xsl:value-of select ="'#95B3D7'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent1' or contains($color,'accent1')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#376092'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent1' or contains($color,'accent1')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#254061'"/>
			</xsl:when>

			<!--Blue, Accent 2-->
			<xsl:when test ="($color = 'accent2' or contains($color,'accent2')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#C0504D'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent2' or contains($color,'accent2')) and $lumMod='20000' and $lumOff='80000'">
				<xsl:value-of select ="'#F2DCDB'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent2' or contains($color,'accent2')) and $lumMod='40000' and $lumOff='60000'">
				<xsl:value-of select ="'#E6B9B8'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent2' or contains($color,'accent2')) and $lumMod='60000' and $lumOff='40000'">
				<xsl:value-of select ="'#D99694'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent2' or contains($color,'accent2')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#953735'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent2' or contains($color,'accent2')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#632523'"/>
			</xsl:when>

			<!--Olive Green Accent 3-->
			<xsl:when test ="($color = 'accent3' or contains($color,'accent3')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#9BBB59'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent3' or contains($color,'accent3')) and $lumMod='20000' and $lumOff='80000'">
				<xsl:value-of select ="'#EBF1DE'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent3' or contains($color,'accent3')) and $lumMod='40000' and $lumOff='60000'">
				<xsl:value-of select ="'#D7E4BD'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent3' or contains($color,'accent3')) and $lumMod='60000' and $lumOff='40000'">
				<xsl:value-of select ="'#C3D69B'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent3' or contains($color,'accent3')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#77933C'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent3' or contains($color,'accent3')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#4F6228'"/>
			</xsl:when>

			<!--Purple Accent 4-->
			<xsl:when test ="($color = 'accent4' or contains($color,'accent4')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#8064A2'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent4' or contains($color,'accent4')) and $lumMod='20000' and $lumOff='80000'">
				<xsl:value-of select ="'#E6E0EC'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent4' or contains($color,'accent4')) and $lumMod='40000' and $lumOff='60000'">
				<xsl:value-of select ="'#CCC1DA'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent4' or contains($color,'accent4')) and $lumMod='60000' and $lumOff='40000'">
				<xsl:value-of select ="'#B3A2C7'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent4' or contains($color,'accent4')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#604A7B'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent4' or contains($color,'accent4')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#403152'"/>
			</xsl:when>

			<!--Aqua Accent 5-->
			<xsl:when test ="($color = 'accent5' or contains($color,'accent5')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#4BACC6'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent5' or contains($color,'accent5')) and $lumMod='20000' and $lumOff='80000'">
				<xsl:value-of select ="'#DBEEF4'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent5' or contains($color,'accent5')) and $lumMod='40000' and $lumOff='60000'">
				<xsl:value-of select ="'#B7DEE8'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent5' or contains($color,'accent5')) and $lumMod='60000' and $lumOff='40000'">
				<xsl:value-of select ="'#93CDDD'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent5' or contains($color,'accent5')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#31859C'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent5' or contains($color,'accent5')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#215968'"/>
			</xsl:when>


			<!--Orange Accent 6-->
			<xsl:when test ="($color = 'accent6' or contains($color,'accent6')) and $lumMod='' and $lumOff=''">
				<xsl:value-of select ="'#F79646'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent6' or contains($color,'accent6')) and $lumMod='20000' and $lumOff='80000'">
				<xsl:value-of select ="'#FDEADA'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent6' or contains($color,'accent6')) and $lumMod='40000' and $lumOff='60000'">
				<xsl:value-of select ="'#FCD5B5'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent6' or contains($color,'accent6')) and $lumMod='60000' and $lumOff='40000'">
				<xsl:value-of select ="'#FAC090'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent6' or contains($color,'accent6')) and $lumMod='75000' and $lumOff=''">
				<xsl:value-of select ="'#E46C0A'"/>
			</xsl:when>
			<xsl:when test ="($color = 'accent6' or contains($color,'accent6')) and $lumMod='50000' and $lumOff=''">
				<xsl:value-of select ="'#984807'"/>
			</xsl:when>

		</xsl:choose >
	</xsl:template >

	<xsl:template name="ConvertEmu">
		<xsl:param name="length" />
		<xsl:param name="unit" />
		<xsl:choose>
			<xsl:when test="$length = '' or not($length) or $length = 0 or format-number($length div 360000, '#.##') = ''">
				<xsl:value-of select="concat(0,'cm')" />
			</xsl:when>
			<xsl:when test="$unit = 'cm'">
				<xsl:value-of select="concat(format-number($length div 360000, '#.##'), 'cm')" />
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="InsertWhiteSpaces">
		<xsl:param name="string" select="."/>
		<xsl:param name="length" select="string-length(.)"/>
		<!-- string which doesn't contain whitespaces-->
		<xsl:choose>
			<xsl:when test="not(contains($string,' '))">
				<xsl:value-of select="$string"/>
			</xsl:when>
			<!-- convert white spaces  -->
			<xsl:otherwise>
				<xsl:variable name="before">
					<xsl:value-of select="substring-before($string,' ')"/>
				</xsl:variable>
				<xsl:variable name="after">
					<xsl:call-template name="CutStartSpaces">
						<xsl:with-param name="cuted">
							<xsl:value-of select="substring-after($string,' ')"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:variable>
				<xsl:if test="$before != '' ">
					<xsl:value-of select="concat($before,' ')"/>
				</xsl:if>
				<!--add remaining whitespaces as text:s if there are any-->
				<xsl:if test="string-length(concat($before,' ', $after)) &lt; $length ">
					<xsl:choose>
						<xsl:when test="($length - string-length(concat($before, $after))) = 1">
							<text:s/>
						</xsl:when>
						<xsl:otherwise>
							<text:s>
								<xsl:attribute name="text:c">
									<xsl:choose>
										<xsl:when test="$before = ''">
											<xsl:value-of select="$length - string-length($after)"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:value-of select="$length - string-length(concat($before,' ', $after))"/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
							</text:s>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
				<!--repeat it for substring which has whitespaces-->
				<xsl:if test="contains($string,' ') and $length &gt; 0">
					<xsl:call-template name="InsertWhiteSpaces">
						<xsl:with-param name="string">
							<xsl:value-of select="$after"/>
						</xsl:with-param>
						<xsl:with-param name="length">
							<xsl:value-of select="string-length($after)"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="CutStartSpaces">
		<xsl:param name="cuted"/>
		<xsl:choose>
			<xsl:when test="starts-with($cuted,' ')">
				<xsl:call-template name="CutStartSpaces">
					<xsl:with-param name="cuted">
						<xsl:value-of select="substring-after($cuted,' ')"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$cuted"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>

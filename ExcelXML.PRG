*-- This PRG file was extracted to a PRG file from the original VCX file by Matt Slay 2017-09-03.
*-- This allows better updates of the source code by the VFP community on GitHib.
*-- GitHub: https://github.com/VFPX/ExcelXML
*---------------------------------------------------------------------------------------
*-- Change Log:
*---------------------------------------------------------------------------------------
*- 2017-09-10:	Ver 1.10
*--				1. Added new method:  ConvertXmlToXlsx(tcFilename, tnFileFormat, tlOpenAfterExporting)
*--				2. Fixed bug in Bottom Border logic if cursor/grid only has 1 row of data.
*--				3. Added new properties for ColumnHeaderBackgroundColor and ColumnHeaderForeColor 
*--				4. Added new property GridClass to use when creating a temporary form to create grid to host the cursor/alias during the export.
*- 
*- 2017-09-03:  Ver 1.09
*--				Added Try/Catch to handle Dynamic properties that do evaluate properly.
*-				[By Matt Slay]
*---------------------------------------------------------------------------------------

Define Class AS_ExcelXML As Custom

	* Array with information about the structure of the table in a ;
	* specified work area, specified by a table alias, or in the currently ;
	* selected work area in an array and returns the number of fields in ;
	* the table.
	* Name of the table/cursor defined in the Grid or name of current ;
	* table/cursor opened.
	Alias           = ''
	* Returns the number of columns included in the Excel file.
	ColumnCount     = 0
	crlf            = ''
	* Specifies the date format.
	DateFormat      = ''
	* Inform the name of Excel file. If you don't inform the name with the ;
	* extension, the XML extension will be included. The default file name ;
	* is "Book1"
	File            = ''
	* Inform the grid control object to convert a grid control in an Excel ;
	* XML file.
	GridObject      = ''
	GridCollection  = '' && (DCA) - 26/01/2023
	* .T. Includes the option Filter in all columns in the generated file.
	HasFilter       = .F.
	Height          = 16
	* .T. locks the header in the generated file. This option in Excel is ;
	* called by Freeze Top Row.
	LockHeader      = .F.
	* .T. to open the file after saving it.
	OpenAfterSaving = .F.
	* Returns the number of rows included in the Excel file.
	RowCount        = 0
	* Defines if the Excel file will have all the grid graphical attributes ;
	* transported.
	SetStyles       = .T.
	* Excel sheet name. The default name is "Sheet1"
	SheetName       = 'Sheet1'
	stylecodenumber = 0
	* Object that contain the information about this class.
	Version         = ''
	Width           = 70
	* XML encoding type used to set the code that defines special ;
	* characters. Default code is "iso-8859-1".
	xmlEncoding     = 'iso-8859-1'
	cErrorMessage = ""
	* The grid class name to use when creating a temporary form to create grid to host the cursor
	* during the export.
	GridClass = "grid"
	* Colmn Header Background color. Can override grid header backcolor. Set to a string with Hex value, like "#CCCCCC" for light gray.
	ColumnHeaderBackgroundColor = .null.
	* Colmn Header ForegColor. Can override grid header forecolor. Set to a string with Hex value, like "#000000" for black.
	ColumnHeaderForeColor = .null.
	
	&& (DCA) - ADDING AS_PROGRESSBAR
	ProgresBar 	= .f.
	Excel 		= .f.
	HasExcel  	= .f.
	StyleCursor = "xxxStylesRowCol"
	
	*|================================================================================ 
	*| ExcelXml::
	Procedure About
	
		Messagebox("ExcelXml " + This.Version.Number + " " + This.Version.Datetime + This.crlf + ;
					"Converts a Grid control into a Microsoft Excel XML file" + This.crlf + ;
					"" + This.crlf + ;
					"Created by " + This.Version.Author + This.crlf + ;
					This.Version.CountryAndCity + This.crlf + ;
					This.Version.url + This.crlf + ;
					This.Version.Email, 64, "About ExcelXml")
	Endproc

	*|================================================================================ 
	*| ExcelXml::
	Procedure AddNewStyle
		Lparameters plcType, plnRow, plnCol, ;
			plcAlignH, plcAlignV, plcFontName, plcFontFamily, ;
			plcFontSize, plcForeColor, plcFontBold, plcFontItalic, ;
			plcFontUnderline, plcFontStrikeThru, plcBackColor, plcPattern, ;
			plcFormat

		Local lcStyleCode, lcXmlStyle
		lcXmlStyle = ""

		*- Definição de bordas entre as linhas/colunas (células)
		lcXmlBorderStyle = ""
		lcTop = "0"
		lcBottom = "0"
		
		If This.GridObject.GridLines >= 1 And This.SetStyles
			lcGridLineWidth = Iif(plcType = "c", Alltrim(Str(Iif(This.GridObject.GridLineWidth >= 4, 3, This.GridObject.GridLineWidth))), "1")
			lcGridLineColor = Iif(plcType = "c", This.ColorToStrHexa(This.GridObject.GridLineColor), This.ColorToStrHexa(Rgb(100, 100, 100)))
			lcXmlBorderStyle = [   <Borders>] + This.crlf

		*- Linhas na horizontal   
			If Inlist(This.GridObject.GridLines, 1, 3)
				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
			EndIf
			
		*- Linhas na vertical
			If Inlist(This.GridObject.GridLines, 2, 3)
				If This.GridObject.GridLines = 2
					If plnRow = 1					&&- Se for a primeira linha
						lcTop = "1"
						lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
					Endif
					If plnRow = (This.RowCount - 1)	&&- Se for a ultima linha
						lcBottom = "1"
						lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
					Endif
				Endif

				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
			Endif

			lcXmlBorderStyle = lcXmlBorderStyle + [   </Borders>] + This.crlf
		Else
			lcXmlBorderStyle = [   <Borders></Borders>] + This.crlf
		Endif


		*- Adiciono no cursor caso nï¿½o ache um registro com os mesmos dados
		If Not Seek(plcAlignH + plcAlignV + plcFontName + plcFontFamily + ;
						plcFontSize + plcForeColor + plcFontBold + plcFontItalic + ;
						plcFontUnderline + plcFontStrikeThru + plcBackColor + plcPattern + ;
						plcFormat + lcTop + lcBottom, ;
						"xxxStylesProperties", "idxStyle")

			This.stylecodenumber = This.stylecodenumber + 1
			lcStyleCode = Alltrim(Lower(plcType)) + Transform(This.stylecodenumber, "@L 99999")

			*- xml de estilo da celula
			lcXmlStyle = [  <Style ss:ID="] + lcStyleCode + [">] + This.crlf + ;
				[   <Alignment ss:Horizontal="] + Alltrim(plcAlignH) + [" ss:Vertical="] + Alltrim(plcAlignV) + ["/>] + This.crlf + ;
				[   <Font ss:FontName="] + Alltrim(plcFontName) + [" x:Family="] + Alltrim(plcFontFamily) + [" ss:Size="] + Alltrim(Str(Val(plcFontSize))) + [" ss:Color="] + Alltrim(plcForeColor) + ["] + This.crlf + ;
				[    ss:Bold="] + plcFontBold + [" ss:Italic="] + plcFontItalic + ["] + Iif(!Empty(plcFontUnderline), [ ss:Underline="] + Alltrim(plcFontUnderline) + ["], "") + [ ss:StrikeThrough="] + Alltrim(plcFontStrikeThru) + ["/>] + This.crlf + ;
				Iif(This.SetStyles, [   <Interior ss:Color="] + Alltrim(plcBackColor) + [" ss:Pattern="] + Alltrim(plcPattern) + ["/>], [   <Interior/>] ) + This.crlf + ;
				Iif(!Empty(plcFormat), [   <NumberFormat ss:Format="] + Alltrim(plcFormat) + ["/>] + This.crlf, "") + ;
				lcXmlBorderStyle + ;
				[  </Style>]

			Insert Into xxxStylesProperties ;
				Values ( 	lcStyleCode, ;
						plcAlignH, ;
						plcAlignV, ;
						plcFontName, ;
						plcFontFamily, ;
						plcFontSize, ;
						plcForeColor, ;
						plcFontBold, ;
						plcFontItalic, ;
						plcFontUnderline, ;
						plcFontStrikeThru, ;
						plcBackColor, ;
						plcPattern, ;
						plcFormat, ;
						lcTop, ;
						lcBottom, ;
						lcXmlStyle )
		Endif

		Insert Into (This.StyleCursor) ;
			Values ( Transform(plnRow, "@L 999999"), ;
					Transform(plnCol, "@L 999"), ;
					xxxStylesProperties.ssCode )

		Return lcXmlStyle
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildColumnsStyles
	
		Local lcAlignH, lcAlignV, lcFontName, lcFontFamily, ;
			lcFontSize, lcForeColor, lcFontBold, lcFontItalic, ;
			lcFontUnderline, lcFontStrikeThru, lcBackColor, lcPattern, ;
			lcFormat, lcXmlBorderStyle, lcXmlStyles, lnRow, lnCol, lnRowFound

		This.stylecodenumber = 0
		lnRow = 0
		lnCol = 0
		lcXmlStyles = ""


		*- Verifico os estilos de todas as linhas/colunas do grid
		Select (This.Alias)
		Go Top
		Scan
			lnRow = lnRow + 1

			If Not This.SetStyles And lnRow >= 2		&&- Nï¿½o aplica os estilos ao grid.
				Exit
			Endif

			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				If Not loColumn.Visible 	&&-Considero somente as colunas visiveis
					Loop
				Endif

				*- Formato dos dados da linha/coluna (cï¿½lula)
				lcDataColumn = Evaluate(loColumn.ControlSource)
				loCurrentControl = This.GetCurrentControlObject(loColumn)
				lcFormat = ""

				If Not Isnull(loCurrentControl)
					Do Case
					Case Inlist(Vartype(lcDataColumn), "N", "Y")
						If Lower(loCurrentControl.BaseClass) $ "textbox//spinner"
							If Not Empty(loColumn.InputMask)
								lcInputMask = loColumn.InputMask
								If Occurs(".", lcInputMask) > 0
									lcFormat = "#,##0." + Replicate("0", Len(Subs(lcInputMask, Rat(".", lcInputMask) + 1)))
								Else
									lcFormat = "#,##0"
								Endif
							Else
								lnRowFound = Ascan(This._Fields, Iif("." $ loColumn.ControlSource, Substr(loColumn.ControlSource, At(".", loColumn.ControlSource) + 1), loColumn.ControlSource), -1, -1, 1, 15)
								If lnRowFound > 0 And This._Fields[lnRowFound, 4] > 0
									lcFormat = "#,##0." + Replicate("0", This._Fields[lnRowFound, 4])
								Else
									lcFormat = ""
								Endif
							Endif
						Endif

						If Lower(loCurrentControl.BaseClass) $ "checkbox//optiongroup"
							lcFormat = ""
						Endif

					Case Vartype(lcDataColumn) = "D"
						lcFormat = This.DateFormat + ";@"

					Case Vartype(lcDataColumn) = "T"
						If Lower(loCurrentControl.BaseClass) = "textbox"
							lnHasSeconds = loCurrentControl.Seconds
						Else
							lnHasSeconds = 2
						Endif

						If lnHasSeconds = 0
							*- Data e hora sem segundos
							lcFormat = This.DateFormat + "\ h:mm" + Iif(Set("hours") = 12, " AM/PM", "")
						Else
							*- Data e hora com segundos
							lcFormat = This.DateFormat + "\ h:mm:ss" + Iif(Set("hours") = 12, " AM/PM", "")
						Endif

					Case Vartype(lcDataColumn) = "L"
						lcFormat = "True/False"

					Otherwise
						lcFormat = ""
					Endcase
				Endif

				lcFormat = Padr(lcFormat, Len(xxxStylesProperties.ssFormat))

				*- Requisitos fixos para o estilo
				lcFontFamily = Padr("Swiss", Len(xxxStylesProperties.ssFontFamily))
				lcPattern = Padr("Solid", Len(xxxStylesProperties.ssPattern))

				*- Alinhamento Horizontal do texto da coluna/linha
				If Not Isnull(loCurrentControl) And Lower(loCurrentControl.BaseClass) = "combobox"
					lcAlignH = This.GetColumnAlign("H", loCurrentControl.Alignment, Vartype(lcDataColumn))
					lcAlignV = This.GetColumnAlign("V", loCurrentControl.Alignment, Vartype(lcDataColumn))
				Else
					lcAlignH = This.GetColumnAlign("H", loColumn.Alignment, Vartype(lcDataColumn))
					lcAlignV = This.GetColumnAlign("V", loColumn.Alignment, Vartype(lcDataColumn))
				Endif

				*- cor de fundo da coluna/linha
				If Not Empty(loColumn.DynamicBackColor)
					Try
						lcBackColor = This.ColorToStrHexa(Evaluate(loColumn.DynamicBackColor) )
					Catch
						lcBackColor = This.ColorToStrHexa( loColumn.BackColor )
					Endtry
				Else
					lcBackColor = This.ColorToStrHexa( loColumn.BackColor )
				Endif


				*- cor da fonte da coluna/linha
				If Not Empty(loColumn.DynamicForeColor)
					Try
						lcForeColor = This.ColorToStrHexa( Evaluate(loColumn.DynamicForeColor) )
					Catch
						lcForeColor = This.ColorToStrHexa( loColumn.ForeColor )
					Endtry
				Else
					lcForeColor = This.ColorToStrHexa( loColumn.ForeColor )
				Endif


				*- fonte usada na coluna/linha
				If Not Empty(loColumn.DynamicFontName)
					Try
						lcFontName = Evaluate(loColumn.DynamicFontName)
					Catch
						lcFontName = Padr(lcFontName, Len(xxxStylesProperties.ssFontName))
					Endtry
				Else
					lcFontName = loColumn.FontName
				Endif
				lcFontName = Padr(lcFontName, Len(xxxStylesProperties.ssFontName))


				*- tamanho da fonte da coluna/linha			
				If Not Empty(loColumn.DynamicFontSize)
					Try
						lcFontSize = Transform(Evaluate(loColumn.DynamicFontSize), "@L 999")
					Catch
						lcFontSize = Transform(loColumn.FontSize, "@L 999")
					Endtry
				Else
					lcFontSize = Transform(loColumn.FontSize, "@L 999")
				Endif


				*- Fonte Italica da coluna/linha
				If Not Empty(loColumn.DynamicFontItalic)
					Try
						lcFontItalic = Iif(Evaluate(loColumn.DynamicFontItalic), "1", "0")
					Catch
						lcFontItalic = Iif(loColumn.FontItalic, "1", "0")
					Endtry
				Else
					lcFontItalic = Iif(loColumn.FontItalic, "1", "0")
				Endif


				*- Fonte Negrito da coluna/linha
				If Not Empty(loColumn.DynamicFontBold)
					Try
						lcFontBold = Iif(Evaluate(loColumn.DynamicFontBold), "1", "0")
					Catch
						lcFontBold = Iif(loColumn.FontBold, "1", "0")
					Endtry
				Else
					lcFontBold = Iif(loColumn.FontBold, "1", "0")
				Endif


				*- Fonte Underline da coluna/linha
				If Not Empty(loColumn.DynamicFontUnderline)
					Try
						lcFontUnderline = Iif(Evaluate(loColumn.DynamicFontUnderline), "Single", "")
					Catch
						lcFontUnderline = Iif(loColumn.FontUnderline, "Single", "")
					Endtry
				Else
					lcFontUnderline = Iif(loColumn.FontUnderline, "Single", "")
				Endif
				lcFontUnderline = Padr(lcFontUnderline, Len(xxxStylesProperties.ssFontUnderline))


				*- Fonte Underline da coluna/linha
				If Not Empty(loColumn.DynamicFontStrikethru)
					Try
						lcFontStrikeThru = Iif(Evaluate(loColumn.DynamicFontStrikethru), "1", "0")
					Catch
						lcFontStrikeThru = Iif(loColumn.FontStrikethru, "1", "0")
					Endtry
				Else
					lcFontStrikeThru = Iif(loColumn.FontStrikethru, "1", "0")
				Endif
				lcFontStrikeThru = Padr(lcFontStrikeThru, Len(xxxStylesProperties.ssFontStrikeThru))


				*- se o estilo já existir "lcXmlStyle" retorna ""
				lcXmlStyle = This.AddNewStyle( This.StyleCursor+":c", lnRow, lnCol, ;
							lcAlignH, lcAlignV, lcFontName, lcFontFamily, ;
							lcFontSize, lcForeColor, lcFontBold, lcFontItalic, ;
							lcFontUnderline, lcFontStrikeThru, lcBackColor, lcPattern, ;
							lcFormat )

				If Not Empty(lcXmlStyle)
					lcXmlStyles = lcXmlStyles + This.crlf + lcXmlStyle
				Endif
			Endfor
		Endscan
		
		Go This._Recno

		Return lcXmlStyles
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildColumnsWidth
	
		Local lcXmlColumnsWidth, lnCol, lnColumnWidth
		
		lcXmlColumnsWidth = This.crlf

		For lnCol = 1 To This.GridObject.ColumnCount
			loColumn = This.GetColumn(lnCol)
			If loColumn.Visible = .T.
				lnColumnWidth = Iif(loColumn.Width > 700, 700, loColumn.Width)		&&- Avoiding error in Excel
				lcXmlColumnsWidth = lcXmlColumnsWidth + [   <Column ss:AutoFitWidth="0" ss:Width="] + Alltrim(Str(lnColumnWidth)) + ["/>] + This.crlf
			Endif
		Endfor

		Return lcXmlColumnsWidth
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildHeadersStyles
	
		Local loColumn, loColumnHeader, lnCol, lcXmlStyles, lcXmlStyle, ;
			lcBackColor, lcForeColor, lcFontName, lcFontSize, lcFontItalic, ;
			lcFontBold,	lcFontUnderline, lcFontStrikeThru, lcFormat, ;
			lcFontFamily, lcPattern, lcAlignH, lcAlignV, lcCollate

		lcXmlStyle = ""
		lcXmlStyles = ""
		This.stylecodenumber = 0
		lcCollate = Set("Collate")

		Set Collate To "MACHINE"

		*- Crio cursor para armazenar todos os estilos encontrados
		Create Cursor xxxStylesProperties ( ssCode c(26), ;
					ssAlignH c(6), ;
					ssAlignV c(6), ;
					ssFontName c(40), ;
					ssFontFamily c(5), ;
					ssFontSize c(3), ;
					ssFontColor c(7), ;
					ssFontBold c(1), ;
					ssFontItalic c(1), ;
					ssFontUnderline c(6), ;
					ssFontStrikeThru c(1), ;
					ssBackColor c(7), ;
					ssPattern c(5), ;
					ssFormat c(40), ;
					ssTop c(1), ;
					ssBottom c(1), ;
					ssStyle m )



		Select xxxStylesProperties
		Index On ssAlignH + ssAlignV + ssFontName +	;
			ssFontFamily + ssFontSize + ssFontColor + ;
			ssFontBold + ssFontItalic + ssFontUnderline + ssFontStrikeThru + ;
			ssBackColor + ssPattern + ssFormat	+ ssTop + ssBottom Tag idxStyle

		Index On ssCode Tag idxCode

		*- Crio cursor para gravar o estilo que sera usado pela linha/coluna (cï¿½lula)
		Create Cursor (This.StyleCursor) ( ssRow c(6), ;
					ssCol c(3), ;
					ssCode c(26) )

		Select (This.StyleCursor)
		Index On ssRow + ssCol Tag idxRowCol

		Set Collate To lcCollate


		*- Verifico os estilos dos headers de cada coluna
		If This.GridObject.HeaderHeight > 0
			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				loColumnHeader = This.GetColumnHeader(loColumn)

				If IsNull(This.ColumnHeaderBackgroundColor)
					lcBackColor = This.ColorToStrHexa( Iif(This.SetStyles, loColumnHeader.BackColor, Rgb(255, 255, 255)) )
				Else
					lcBackColor = This.ColumnHeaderBackgroundColor
				EndIf
				
				If IsNull(This.ColumnHeaderForeColor)
					lcForeColor = This.ColorToStrHexa( Iif(This.SetStyles, loColumnHeader.ForeColor, Rgb(0, 0, 0)) )
				Else
					lcForeColor = This.ColumnHeaderForeColor
				EndIf
				
				lcFontName = Padr(loColumnHeader.FontName, Len(xxxStylesProperties.ssFontName))
				lcFontSize = Transform(loColumnHeader.FontSize, "@L 999")
				lcFontItalic = Iif(loColumnHeader.FontItalic, "1", "0")
				lcFontBold = Iif(loColumnHeader.FontBold Or This.SetStyles = .F., "1", "0")
				lcFontUnderline = Padr(Iif(loColumnHeader.FontUnderline Or This.SetStyles = .F., "Single", ""), Len(xxxStylesProperties.ssFontUnderline))
				lcFontStrikeThru = Iif(loColumnHeader.FontStrikethru, "1", "0")
				lcFormat = Padr("", Len(xxxStylesProperties.ssFormat))
				lcFontFamily = Padr("Swiss", Len(xxxStylesProperties.ssFontFamily))
				lcPattern = Padr("Solid", Len(xxxStylesProperties.ssPattern))
				lcAlignH = Iif(This.SetStyles, This.GetColumnAlign("H", loColumnHeader.Alignment), "Left")
				lcAlignV = Iif(This.SetStyles, This.GetColumnAlign("V", loColumnHeader.Alignment), "Center")

				*- se o estilo jï¿½ existir "lcXmlStyle" retorna ""
				lcXmlStyle = This.addnewstyle( This.StyleCursor+":h", 0, lnCol, ;
							lcAlignH, lcAlignV, lcFontName, lcFontFamily, ;
							lcFontSize, lcForeColor, lcFontBold, lcFontItalic, ;
							lcFontUnderline, lcFontStrikeThru, lcBackColor, lcPattern, ;
							lcFormat )

				If Not Empty(lcXmlStyle)
					lcXmlStyles = lcXmlStyles + This.crlf + lcXmlStyle
				Endif
			Endfor
		Endif

		Return lcXmlStyles
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildRows
		Local lcXmlRows, lcDataType, lcDataColumn, lcAuxDataColumn, lnRow, lnCol, loColumn, loColumnHeader, loCurrentControl, ;
			lnPercent, lnCountRowSource, lcCountOption, laArrayTmp, lcComboOption, lcToolTipText, lnBytes, llHasDecimals, ;
			lnSetDecimals, lnRowFound, lnYear

		lcXmlRows = This.crlf
		lnRow = 0
		lnCol = 0
		lnBytes = 0
		lnSetDecimals = Set("Decimals")
		
		*- Adiciono a linha do Header no arquivo excel
		If This.GridObject.HeaderHeight > 0
			lcXmlRows = lcXmlRows + [   <Row ss:AutoFitHeight="0" ss:Height="] + Alltrim(Str(This.GridObject.HeaderHeight)) + [">] + This.crlf
			
			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				loColumnHeader = This.GetColumnHeader(loColumn)

				If loColumn.Visible = .T.
					*- caso tenha tooltiptext
					lcToolTipText = ""
					If Not Empty(loColumnHeader.ToolTipText)
						lcToolTipText = [<Comment ss:Author="Rodrigo_Bruscain">] + ;
							[<ss:Data xmlns="http://www.w3.org/TR/REC-html40">] + ;
							[<Font html:Face="Tahoma" x:Family="Swiss" html:Color="#000000">] + Alltrim(loColumnHeader.ToolTipText) + [</Font>] + ;
							[</ss:Data>] + ;
							[</Comment>]
					Endif

					*- linha do header
					lcXmlRows = lcXmlRows + [     <Cell ss:StyleID="] + This.SeekStyle("000000", Transform(lnCol, "@L 999")) + ["><Data ss:Type="String">] + loColumnHeader.Caption + [</Data>] + lcToolTipText + [</Cell>] + This.crlf
				Endif
			Endfor

			lcXmlRows = lcXmlRows + [   </Row>] + This.crlf
		Endif

		lcXmlRows = lcXmlRows + This.crlf
		
		*- Adiciono a linha do Registro no arquivo excel
		Select (This.Alias)
		Go Top
		
		Scan
			lnRow = lnRow + 1
			lcXmlRows = lcXmlRows + [   <Row ss:AutoFitHeight="0">] + This.crlf

			*- percentual processado
			lnPercent = Int((lnRow / (This.RowCount - (Iif(This.GridObject.HeaderHeight > 0, 1, 0))) ) * 100)
			This.Progress(lnPercent)

			*- faï¿½o a varredura em todas as colunas
			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				If Not loColumn.Visible
					Loop
				Endif

				*- Verifico o tipo de dado da coluna
				lcDataColumn = Evaluate(loColumn.ControlSource)
				loCurrentControl = This.GetCurrentControlObject(loColumn)

				*- se nï¿½o tem objeto de controle na linha da coluna nï¿½o levo a informaï¿½ï¿½o da tabela ao excel
				If Isnull(loCurrentControl)
					lcDataType = "String"
					lcDataColumn = ""
				Else
					Do Case
					Case Vartype(lcDataColumn) $ "N//Y"
						lcDataType = "Number"

						*- Se o currentcontrol da coluna for um combobox mostro o seu conteudo ao inves da posiï¿½ï¿½o numerica
						If Lower(loCurrentControl.BaseClass) = "combobox"
							Try
								Do Case
								*- Mostro o texto do value
								Case loCurrentControl.RowSourceType = 1
									lcDataType = "String"

									If Not Empty(loCurrentControl.RowSource)
										lcAuxDataColumn = Alltrim(loCurrentControl.RowSource)
										lcAuxDataColumn = Strtran(Strtran(Strtran(lcAuxDataColumn, " ,", ","), ", ", ","), " , ", ",")
										lcCountOption = Occurs(",", lcAuxDataColumn) + 1

										Dimension laArrayTmp[lcCountOption]
										For lnCountRowSource = 1 To lcCountOption
											lcComboOption = Substr(lcAuxDataColumn, 1, Iif(lnCountRowSource < lcCountOption, At(",", lcAuxDataColumn) - 1, Len(lcAuxDataColumn)) )
											lcAuxDataColumn = Strtran(lcAuxDataColumn, lcComboOption + Iif(lcCountOption >= 2, ",", ""), "")
											laArrayTmp[lnCountRowSource] =  lcComboOption
										Endfor

										lcDataColumn = Evaluate("laArrayTmp[" + Alltrim(Str(lcDataColumn)) + "]")
									Endif

								*- Mostro o texto do array do combo	
								Case loCurrentControl.RowSourceType = 5
									lcDataType = "String"

									*- Se for um array objeto ex: thisform.ArrayName ou MyObj.ArrayName
									If Occurs(".", loCurrentControl.RowSource) > 0
										lcObjArrayName = Substr(loCurrentControl.RowSource, 1, Rat(".", loCurrentControl.RowSource) - 1)

										*- Se for um array objeto publico
										If Type(lcObjArrayName) = "O"
											lcAuxDataColumn = loCurrentControl.RowSource + "[" + Alltrim(Str(lcDataColumn)) + "]"
										Else
											lcArrayName = Substr(loCurrentControl.RowSource, Rat(".", loCurrentControl.RowSource) + 1)
											lnCountObjectHierarchy = Occurs(".", Sys(1272, This.GridObject))
											lcAuxDataColumn = "This.GridObject" + Replicate(".Parent", lnCountObjectHierarchy) + "." + lcArrayName + "[" + Alltrim(Str(lcDataColumn)) + "]"
										Endif

									*- Array comum	
									Else
										lcAuxDataColumn = loCurrentControl.RowSource + "[" + Alltrim(Str(lcDataColumn)) + "]"
									Endif

									lcDataColumn = Evaluate(lcAuxDataColumn)


								*- Qualquer outro mostro o conteudo do campo e nï¿½o o conteudo do array
								Otherwise
									lcDataColumn = lcDataColumn
								Endcase

							Catch To loError
							Endtry

							If Vartype(loError) = "O"
								Messagebox( "Combo array '" + loCurrentControl.RowSource + "' in column '" + loColumn.Name + "' not is valid", 48)
								Select (This.Alias)
								Go Top
								Return .F.
							Endif
						Else

							lnRowFound = Ascan(This._Fields, Iif("." $ loColumn.ControlSource, Substr(loColumn.ControlSource, At(".", loColumn.ControlSource) + 1), loColumn.ControlSource), -1, -1, 1, 15)
							If lnRowFound > 0 And This._Fields[lnRowFound, 4] > 0
								llHasDecimals = .T.
								Set Decimals To This._Fields[lnRowFound, 4]
							Else
								llHasDecimals = .F.
							Endif

						Endif

					Case Vartype(lcDataColumn) = "D"
						lcDataType = "DateTime"
						If Not Empty(Nvl(lcDataColumn, ""))
							lnYear = Iif(Year(lcDataColumn) < 1900, 1900, Year(lcDataColumn))
							lcAuxDataColumn = Str(lnYear, 4) + "-" + Transform(Month(lcDataColumn), "@L 99") + "-" + Transform(Day(lcDataColumn), "@L 99") + "T00:00:00.000"
							lcDataColumn = lcAuxDataColumn
						Else
							lcDataType = "String"
							lcDataColumn = ""
						Endif

					Case Vartype(lcDataColumn) = "T"
						lcDataType = "DateTime"
						If Not Empty(Nvl(lcDataColumn, ""))
							lnYear = Iif(Year(lcDataColumn) < 1900, 1900, Year(lcDataColumn))
							lcAuxDataColumn = Str(lnYear, 4) + "-" + Transform(Month(lcDataColumn), "@L 99") + "-" + Transform(Day(lcDataColumn), "@L 99") + ;
								"T" + Transform(Hour(lcDataColumn), "@L 99") + ":" + Transform(Minute(lcDataColumn), "@L 99") + ":" + Transform(Sec(lcDataColumn), "@L 99") + ".000"
							lcDataColumn = lcAuxDataColumn
						Else
							lcDataType = "String"
							lcDataColumn = ""
						Endif

					Case Vartype(lcDataColumn) = "L"
						lcDataType = "Number"
						lcDataColumn = Iif(lcDataColumn, 1, 0)

					Otherwise
						lcDataType = "String"
						If Isnull(lcDataColumn)
							lcDataColumn = ""
						Endif
					Endcase
				Endif

				*- removing invalid characters
				If lcDataType = "String" And ("<" $ lcDataColumn Or ">" $ lcDataColumn)
					lcDataColumn = Strtran(Strtran(lcDataColumn, "<", "["), ">", "]")
				Endif

				*- incluo a linha de dados
				lcXmlRows = lcXmlRows + [     <Cell ss:StyleID="] + This.SeekStyle(Transform(lnRow, "@L 999999"), Transform(lnCol, "@L 999")) + ["><Data ss:Type="] + lcDataType + [">] + Alltrim(Transform(lcDataColumn, "")) + [</Data></Cell>] + This.crlf

				*- devolvo o atributo original
				If llHasDecimals
					Set Decimals To lnSetDecimals
				Endif
			Endfor

			lcXmlRows = lcXmlRows + [   </Row>] + This.crlf
			lnBytes = lnBytes + Strtofile( lcXmlRows + This.crlf, This.File, 1)
			lcXmlRows = ""

		Endscan
		
		Go This._Recno
		
		Return lnBytes
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure ColorToStrHexa(plnColor)
	
		Local lnDecimalColor
	
		lnDecimalColor = Substr(Transform(plnColor, '@0'), 5)
		Return "#" + Right(lnDecimalColor, 2) + Substr(lnDecimalColor, 3, 2) + Left(lnDecimalColor, 2)
	
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetColumn(plcColumnNumber)

		Local lnCol

		For lnCol = 1 To This.GridObject.ColumnCount
			If This.GridObject.Columns(lnCol).ColumnOrder = plcColumnNumber
				Return This.GridObject.Columns(lnCol)
			Endif
		EndFor
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetColumnAlign(plcWhat, plnAlignment, plcVartype)
	
		Local lcAlignment, lcAlignH, lcAlignV
		
		plcVartype = Evl(plcVartype, "")
		lcAlignment = Alltrim(Str(plnAlignment))

		*- Alinhamento Horizontal do texto da coluna/linha
		If plcWhat = "H"
			Do Case
			Case lcAlignment $ "0//4//7"
				lcAlignH = "Left"
			Case lcAlignment $ "1//5//8"
				lcAlignH = "Right"
			Case lcAlignment $ "2//6//9"
				lcAlignH = "Center"
			Otherwise
				lcAlignH = Iif(plcVartype $ "N//Y", "Right", "Left")
			Endcase

			lcAlignH = Padr(lcAlignH, Len(xxxStylesProperties.ssAlignH))
			Return lcAlignH
		Endif

		*- Alinhamento vertical do texto da coluna/linha
		If plcWhat = "V"
			Do Case
			Case lcAlignment $ "4//5//6"
				lcAlignV = "Top"
			Case lcAlignment $ "7//8//9"
				lcAlignV = "Bottom"
			Case lcAlignment $ "0//1//2"
				lcAlignV = "Center"
			Otherwise
				lcAlignV = "Center"
			Endcase

			lcAlignV = Padr(lcAlignV, Len(xxxStylesProperties.ssAlignV))
			Return lcAlignV
		EndIf
	
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetColumnHeader(ploColumn)
		
		Local loReturn, lnX
		loReturn = ""

		If ploColumn.ControlCount > 0
			For lnX = 1 To ploColumn.ControlCount
				If Lower(ploColumn.Controls(lnX).BaseClass) = "header"
					loReturn = ploColumn.Controls(lnX)
					Exit
				Endif
			Endfor
		Endif

		Return loReturn
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetCurrentControlObject(ploGridColumn)
		
		Local lcCurrentControl

		If Not Empty(ploGridColumn.DynamicCurrentControl)
			Try
				lcCurrentControl = Evaluate(ploGridColumn.DynamicCurrentControl)
			Catch
				lcCurrentControl = ploGridColumn.CurrentControl
			Endtry
		Else
			lcCurrentControl = ploGridColumn.CurrentControl
		Endif

		If Not Empty(lcCurrentControl)
			Return Evaluate("ploGridColumn." + lcCurrentControl)
		Else
			Return Null
		EndIf
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure HasColumnVisible
	
		Local lnCol, llReturn
		llReturn = .F.

		For lnCol = 1 To This.GridObject.ColumnCount
			If This.GridObject.Columns(lnCol).Visible
				llReturn = .T.
				Exit
			Endif
		Endfor

		Return llReturn
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	*/---------------------------------------------------------------------------------------------------/*
	*/ Descripton..: - Classe para converter o grid do vfp em um arquivo xml para o Excel.               /*
	*/				 - A grande vantagem na utilizaï¿½ï¿½o ï¿½ que Nï¿½O NECESSITA DO EXCEL INSTALADO            /*
	*/                 pois em nenhum momento o Excel ï¿½ instanciado para automaï¿½ï¿½o.                      /*
	*/                 Apesar de ser um arquivo xml, se encontra no padrï¿½o Microsoft onde ï¿½ reconhecido  /*
	*/                 pelo Excel como "Planilha XML 2003 (*.xml)". Dessa forma fica restrito o uso      /*
	*/                 para Excel 2003 ou superior.                                                      /*
	*/                                                                                                   /*
	*/				 - Se o Excel estiver instalado o icone do arquivo gerado serï¿½ reconhecido		     /*
	*/				   pelo Excel e abrindo o arquivo serï¿½ reconhecido como se fosse um XLS ou XLSX,     /*		
	*/                 ou seja, tudo serï¿½ transparente para o Excel.                                     /*
	*/                                                                                                   /*
	*/               - Praticamente todos os recursos visuais do grid, headers, colunas e linhas         /*
	*/                 sï¿½o tratados na exportaï¿½ï¿½o. Segue abaixo as propriedades reconhecidas:            /*
	*/                                                                                                   /*
	*/                 Header Properties                                                                 /*
	*/                 ---------------------------------                                                 /*  
	*/                 ToolTipText / HeaderHeight / Alignment / FontBold / FontItalic / FontUnderline /  /*
	*/                 FontStrikeThru  / FontName / FontSize / ForeColor / BackColor / Caption /         /*
	*/                                                                                                   /*
	*/                 Columns Properties                                                                /*
	*/                 ---------------------------------                                                 /*  
	*/				   ControlSource / BaseClass / InputMask / Seconds / RowHeight / Alignment / 		 /*
	*/				   FontBold / FontItalic / FontUnderline / FontStrikeThru  / FontName / FontSize /   /*
	*/				   ForeColor / FontBackColor / CurrentControl / DynamicFontBold / DynamicFontItalic  /*
	*/				   DynamicFontUnderline / DynamicFontStrikeThru / DynamicCurrentControl /         	 /*
	*/				   DynamicFontName / DynamicFontSize / DynamicForeColor / DynamicBackColor / 		 /*	
	*/				   ColumnCount / ColumnOrder / Width / Visible / Combobox.Alignment /                /*
	*/                 Combobox.RowSource / Combobox.RowSourceType					 					 /*
	*/                																					 /*
	*/                 Environment																		 /*			
	*/                 ---------------------------------                                                 /*  
	*/				   set date / set century / set hours                                                /*
	*/                                                                                                   /*
	*/
	*/				   Goals
	*/                 ------
	*/                 a) Possibilidade de gerar planilhas com mais de 65,535 linhas superando 
	*/                    a limitaï¿½ao nativa do VFP
	*/                 b) Converte um grid em planilha Excel assumindo 99% do visual do grid
	*/				   c) Easy to implement and it is not necessary to change your code
	*/                 d) Compativel com Excel 2003 ou superior
	*/                 e) Pode ser aberto pelo OpenOffice reduzindo erros de conversï¿½o
	*/                 f) Ao abrir o arquivo pelo Excel ï¿½ possivel salvar em outros formatos
	*/                 g) Nao precisa ter o Excel instalado
	*/
	*/                                                                                                   /*
	*/ Original Author......: Rodrigo Bruscain                                                           /*
	*/ Original Date........: 25/05/2013 (Original)                                                      /*
	*/ Country.....: Brazil - Sï¿½o Paulo - SP                                                             /*
	*/---------------------------------------------------------------------------------------------------/*
	Procedure Init

		This.crlf = Chr(13) + Chr(10)
		
		Local lcDateFormat, lcCentury

		AddProperty(This, "_Fields[1]")
		AddProperty(This, "_Recno")
		Dimension This._Fields[1,18]
		
		lcDateFormat = Set("Date")
		lcCentury = Iif(Set("century") = "ON", "yyyy", "yy")

		Do Case
		Case Inlist(lcDateFormat, "AMERICAN", "MDY")				&& month/day/year
			This.DateFormat = "mm/dd/" + lcCentury

		Case lcDateFormat = "ANSI"									&& year.month.day
			This.DateFormat = lcCentury + ".mm.dd"

		Case Inlist(lcDateFormat, "BRITISH", "DMY", "FRENCH") 		&& day/month/year
			This.DateFormat = "dd/mm/" + lcCentury

		Case lcDateFormat = "GERMAN"								&& day.month.year
			This.DateFormat = "dd.mm." + lcCentury

		Case lcDateFormat = "ITALIAN"								&& day-month-year
			This.DateFormat = "dd-mm-" + lcCentury

		Case Inlist(lcDateFormat, "JAPAN", "YMD")					&& year/month/day
			This.DateFormat = lcCentury + "/mm/dd"

		Case lcDateFormat = "USA"									&& month-day-year
			This.DateFormat = "mm-dd-" + lcCentury

		Otherwise
			This.DateFormat = "dd/mm/" + lcCentury
		Endcase

		*- version object
		This.Version = Createobject("empty")
		AddProperty(This.Version, "Version", "1.10")
		AddProperty(This.Version, "DateTime", "Sep.10.2017 3:59:41 AM")
		AddProperty(This.Version, "Author", "Rodrigo Duarte Bruscain")
		AddProperty(This.Version, "CountryAndCity", "kitchener ON - Canada")
		AddProperty(This.Version, "Url", "https://github.com/ExcelXml")
		AddProperty(This.Version, "Email", "bruscain@hotmail.com")
		AddProperty(This.Version, "Email2", "mattslay@jordanmachine.com")
		
		do FORM AS_PROGRESSBAR name AS_PROGRESSBAR noshow
		This.ProgresBar = AS_PROGRESSBAR
		This.ProgresBar.show(0)
		Release AS_PROGRESSBAR
		
		Try 
			This.Excel 		= Createobject("Excel.Application")
			This.hasExcel	= .T.
		Catch to oError 
			This.hasExcel	= .F.
		EndTry
		
	Endproc
	
	Procedure Destroy 
		If Type("This.Excel") == "O"
			This.Excel.Quit()
		Endif
	Endproc
	
	*|================================================================================ 
	*| ExcelXml::
	Procedure Progress(plnPercent)
	
		*-- Add any code here that you want to execute as processing scans over each row...
		This.ProgresBar.as_progressbar.Update(plnPercent)
		
	*	If plnPercent == 100
	*		wait Window "" TimeOut 1
	*		This.ProgresBar.Release()	
	*	Endif
		
	Endproc
	
	*** <summary>
	*** Adiciona o Grid a Collection de Grids
	*** </summary>
	*** <param name="plo_Grid">Grid a ser exportado</param>
	*** <param name="plc_SheetName">Nome da aba do Excel</param>
	*** <remarks></remarks>
	Procedure AddGrid
		Lparameters plo_Grid, plc_SheetName
		
		Local lnCount
		
		If Type('This.GridCollection') <> 'O'
			This.GridCollection = CreateObject("Collection")	
		EndIf
		
		*- crio um grid virtual caso a nao exista um grid para conversao, 
		*- ou seja, estou convertendo somente a tabela
		If VarType(plo_Grid) != "O"
			loForm = CreateObject("form")
			loForm.AddObject("grid1", This.GridClass)
			loForm.Grid1.RecordSource =  Alias()
			loForm.Grid1.Visible = .T.
			loForm.Refresh()
			plo_Grid = loForm.Grid1
			This.SetStyles = .F.
		Endif
	
		This.GridCollection.add(CreateObject("EMPTY")) 
		lnCount = This.GridCollection.Count
		
		plc_SheetName = Iif(Type('plc_SheetName') == 'C', Alltrim(plc_SheetName), "Sheet"+Transform(This.GridCollection.Count))
		AddProperty(This.GridCollection[lnCount],"Grid", plo_Grid)
		AddProperty(This.GridCollection[lnCount],"SheetName", plc_SheetName)
		
	Endproc
		
	*** <summary>
	*** Build Grid Object
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildGridObject
		Lparameters pln_GridIndex
		Local loForm, lnCol
	
		*- crio um grid virtual caso a nao exista um grid para conversao, 
		*- ou seja, estou convertendo somente a tabela
*!*			If VarType(This.GridObject) != "O"
*!*				loForm = CreateObject("form")
*!*				loForm.AddObject("grid1", This.GridClass)
*!*				loForm.Grid1.RecordSource =  Alias()
*!*				loForm.Grid1.Visible = .T.
*!*				loForm.Refresh()
*!*				This.GridObject = loForm.Grid1
*!*				This.SetStyles = .F.
*!*			Endif

		pln_GridIndex = Int(Val(Transform(pln_GridIndex)))
		If pln_GridIndex == 0
			Return .f.
		Endif
		
		This.GridObject 	= This.GridCollection[pln_GridIndex].Grid
		This.SheetName  	= This.GridCollection[pln_GridIndex].SheetName
		This.StyleCursor	= "xxxStylesRowCol" + Padl(pln_GridIndex, 4,"0")
		
		*- environment
		If This.GridObject.RecordSourceType = 1
			This.Alias = This.GridObject.RecordSource
		Else
			This.Alias = Alias()
		Endif

		This._Recno = Recno(This.Alias)
		Afields(This._Fields, This.Alias)

		*- Numero de colunas válidas para o excel
		This.ColumnCount = 0
		For lnCol = 1 To This.GridObject.ColumnCount
			If This.GridObject.Columns(lnCol).Visible = .T.
				This.ColumnCount = This.ColumnCount + 1
			Endif
		Endfor

		*- Numero de linhas disponï¿½veis para o excel
		This.RowCount = 0
		Select (This.Alias)
		Count To This.RowCount
		Go Top

		If This.GridObject.HeaderHeight > 0
			This.RowCount = This.RowCount + 1
		Endif

		*- verifico se tudo esta ok para prosseguir
		If Isnull(This.GridObject) Or This.GridObject.ColumnCount <= 0 And This.hascolumnvisible()
			Return .F.
		Endif
	Endproc
	
	*** <summary>
	*** Build Xml Start
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildXmlStart
		Local lcXmlStart
		*- Inicio tratamento dos dados
		Text To lcXmlStart Textmerge Pretext 2 Noshow
			<?xml version="1.0" encoding="<<This.xmlEncoding>>"?>
			<?mso-application progid="Excel.Sheet"?>
			<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
			 xmlns:o="urn:schemas-microsoft-com:office:office"
			 xmlns:x="urn:schemas-microsoft-com:office:excel"
			 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
			 xmlns:html="http://www.w3.org/TR/REC-html40">
		EndText
		
		Return lcXmlStart
	Endproc	

	*** <summary>
	*** Build Document Properties
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildDocumentProperties
		Local lcXmlDocumentProperties, lcCreatedDate
		
		*- Data da criação do arquivo excel
		lcCreatedDate = Str(Year(Date()), 4) + "-" + Transform(Month(Date()), "@L 99") + "-" + Transform(Day(Date()), "@L 99") + "T" + Time() + "Z"

		
		Text To lcXmlDocumentProperties Textmerge Pretext 2 Noshow
			 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
			  <Author><<iif(!empty(getenv("USERNAME")), getenv("USERNAME"), iif(!empty(getenv("COMPUTERNAME")), getenv("COMPUTERNAME"), "RODRIGO_BRUSCAIN"))>></Author>
			  <LastAuthor><<iif(!empty(getenv("USERNAME")), getenv("USERNAME"), iif(!empty(getenv("COMPUTERNAME")), getenv("COMPUTERNAME"), "RODRIGO_BRUSCAIN"))>></LastAuthor>
			  <Created><<lcCreatedDate>></Created>
			  <LastSaved><<lcCreatedDate>></LastSaved>
			  <Version>12.00</Version>
			 </DocumentProperties>
		ENDTEXT
		
		Return lcXmlDocumentProperties
	Endproc

	*** <summary>
	*** Build XmlExcel Workbook
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildXmlExcelWorkbookHeader()
		
		Local lcXmlExcelWorkbook
		Text To lcXmlExcelWorkbook Textmerge Pretext 2 Noshow
			 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
			  <WindowHeight>8130</WindowHeight>
			  <WindowWidth>15135</WindowWidth>
			  <WindowTopX>120</WindowTopX>
			  <WindowTopY>45</WindowTopY>
			  <ProtectStructure>False</ProtectStructure>
			  <ProtectWindows>False</ProtectWindows>
			 </ExcelWorkbook>
		EndText
		
		Return lcXmlExcelWorkbook
	Endproc

	*** <summary>
	*** Build All Styles
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildAllStyles
		
		Local lcStringStyles, lcXmlAllStyles, lnx, lcAlias, lcSelectAlias
		lcAlias = Alias()
		lcSelectAlias = Iif(!Empty(lcAlias), "Select "+lcAlias, "")
		
		*- Crio os estilos de cores/fontes/formato/etc das colunas
		*- Depois junto com o estilo padrï¿½o todos os estilos encontrados
		*- Estilos sï¿½o todas as formatï¿½es da cï¿½lulas combinadas onde um estilo pode ser usado
		*- por vï¿½rias cï¿½luas ou por uma ï¿½nica cï¿½lula.
		
		lnx = 1
		lcXmlAllStyles = ""
		lcStringStyles = ""

		For lnx = 1 to This.GridCollection.Count
			
			If lnx > 1 and !This.BuildGridObject(lnx)
				&lcSelectAlias
				Return lcXmlAllStyles
			EndIf
			
			lcStringStyles = lcStringStyles + This.BuildHeadersStyles()	&&- Estilos do header
			lcStringStyles = lcStringStyles + This.buildcolumnsstyles()	&&- Estilos das linhas/colunas
			
			If Used(This.Alias) AND Reccount(This.Alias) > This._Recno
				Select (This.Alias)
				Go This._Recno
			EndIf
			
		EndFor
		
		If !This.BuildGridObject(1)
			&lcSelectAlias
			Return lcXmlAllStyles
		EndIf
	
		Text To lcXmlAllStyles Textmerge Pretext 2 Noshow
			 <Styles>
			  <Style ss:ID="Default" ss:Name="Normal">
			   <Alignment ss:Vertical="Center"/>
			   <Borders/>
			   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"/>
			   <Interior/>
			   <NumberFormat/>
			   <Protection/>
			  </Style>
			  <<lcStringStyles>>
			 </Styles>
		ENDTEXT
		
		&lcSelectAlias
		
		Return lcXmlAllStyles
	Endproc
	
	*** <summary>
	*** Build Xml Freeze Panes
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildXmlFreezePanes
		
		Local lcXmlFreezePanes
		
		*- Congelando paineis na horizontal e vertical 
		Do Case
		*- Congelo a linha do header
		Case This.GridObject.LockColumns = 0 And (This.GridObject.HeaderHeight > 0 And This.LockHeader)
			Text To lcXmlFreezePanes Textmerge Pretext 2 Noshow
				   <FreezePanes/>
				   <FrozenNoSplit/>
				   <SplitHorizontal>1</SplitHorizontal>
				   <TopRowBottomPane>1</TopRowBottomPane>
				   <ActivePane>2</ActivePane>
				   <Panes>
				    <Pane>
				     <Number>3</Number>
				    </Pane>
				    <Pane>
				     <Number>2</Number>
				    </Pane>
				   </Panes>
			ENDTEXT

		*- congelo a linha do header e a coluna definida 
		Case This.GridObject.LockColumns > 0 And (This.GridObject.HeaderHeight > 0 And This.LockHeader)
			Text To lcXmlFreezePanes Textmerge Pretext 2 Noshow
				   <FreezePanes/>
				   <FrozenNoSplit/>
				   <SplitHorizontal>1</SplitHorizontal>
				   <TopRowBottomPane>1</TopRowBottomPane>
				   <SplitVertical><<alltrim(str(This.GridObject.LockColumns))>></SplitVertical>
				   <LeftColumnRightPane><<alltrim(str(This.GridObject.LockColumns))>></LeftColumnRightPane>
				   <ActivePane>0</ActivePane>
				   <Panes>
				    <Pane>
				     <Number>3</Number>
				    </Pane>
				    <Pane>
				     <Number>1</Number>
				    </Pane>
				    <Pane>
				     <Number>2</Number>
				    </Pane>
				    <Pane>
				     <Number>0</Number>
				    </Pane>
				   </Panes>			 
			ENDTEXT

		*- congelo somente a coluna definida
		Case This.GridObject.LockColumns > 0 And (This.GridObject.HeaderHeight = 0 Or Not This.LockHeader)
			Text To lcXmlFreezePanes Textmerge Pretext 2 Noshow
				   <FreezePanes/>
				   <FrozenNoSplit/>
				   <SplitVertical>2</SplitVertical>
				   <LeftColumnRightPane>2</LeftColumnRightPane>
				   <ActivePane>1</ActivePane>
				   <Panes>
				    <Pane>
				     <Number>3</Number>
				    </Pane>
				    <Pane>
				     <Number>1</Number>
				    </Pane>
				   </Panes>
			ENDTEXT

		Otherwise
			lcXmlFreezePanes = ""
		Endcase

		Return lcXmlFreezePanes
	Endproc

	*** <summary>
	*** Build Worksheet Header
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildWorksheetHeader
		
		Local lcXmlWorksheet_part1, lcStringColumnWidth
		
		*- tratamento do nome da planilha
		This.SheetName = Chrtran(Alltrim(Substr(This.SheetName, 1, 31)), ':?][*/\', '')
		This.SheetName = Iif(Empty(This.SheetName), "Sheet1", This.SheetName)

		*- Monto a tabela
		lcStringColumnWidth = This.buildcolumnswidth()

		Text To lcXmlWorksheet_part1 Textmerge Pretext 2 Noshow
			 <Worksheet ss:Name="<<This.SheetName>>">
			  <Table ss:ExpandedColumnCount="<<This.ColumnCount>>" ss:ExpandedRowCount="<<This.RowCount>>" x:FullColumns="1"
			   x:FullRows="1" ss:DefaultRowHeight="<<alltrim(str(This.GridObject.RowHeight-3))>>">
			   <<lcStringColumnWidth>>
		EndText
		
		Return lcXmlWorksheet_part1
	Endproc

	*** <summary>
	*** Build Worksheet Footer
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildWorksheetFooter
		
		Local lcXmlWorksheet_part2, lcStringFilter, lcXmlFreezePanes

		*- Congelando paineis na horizontal e vertical 
		lcXmlFreezePanes = This.BuildXmlFreezePanes()
	
		lcStringFilter = ""
		If This.HasFilter And This.GridObject.HeaderHeight > 0
			Text To lcStringFilter Textmerge Pretext 2 Noshow
				<AutoFilter x:Range="R1C1:R<<alltrim(str(This.RowCount))>>C<<alltrim(str(This.ColumnCount))>>"
				 xmlns="urn:schemas-microsoft-com:office:excel">
				</AutoFilter>
			ENDTEXT
		Endif

		Text To lcXmlWorksheet_part2 Textmerge Pretext 2 Noshow
			  </Table>
			  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
			   <PageSetup>
			    <Header x:Margin="0.31496062000000002"/>
			    <Footer x:Margin="0.31496062000000002"/>
			    <PageMargins x:Bottom="0.78740157499999996" x:Left="0.511811024"
			     x:Right="0.511811024" x:Top="0.78740157499999996"/>
			   </PageSetup>
			   <Unsynced/>
			   <Print>
			    <ValidPrinterInfo/>
			    <PaperSizeIndex>9</PaperSizeIndex>
			    <HorizontalResolution>300</HorizontalResolution>
			    <VerticalResolution>300</VerticalResolution>
			   </Print>
			   <Selected/>
			   <<lcXmlFreezePanes>>
			   <ProtectObjects>False</ProtectObjects>
			   <ProtectScenarios>False</ProtectScenarios>
			  </WorksheetOptions>
			  <<lcStringFilter>> 	
			 </Worksheet>
		EndText
		
		Return lcXmlWorksheet_part2
	Endproc

	*** <summary>
	*** Build Xml Excel Workbook Footer
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure BuildXmlExcelWorkbookFooter
		Local lcXmlExcelWorkbookFooter
		
		TEXT TO XmlExcelWorkbookFooter Textmerge Pretext 2 Noshow
			</Workbook>
		ENDTEXT
		
		Return XmlExcelWorkbookFooter
	Endproc

	*|================================================================================ 
	*| ExcelXml::
	Procedure Save(plcFile)
		
		*local lcAlias, lnRecNo && Not being used
		Local lcSetPoint, lnBytes, loError, lcAlias, lcSelectAlias
		lcAlias = Alias()
		lcSelectAlias = Iif(!Empty(lcAlias), "Select "+lcAlias, "")
		
		plcFile = Evl(plcFile, "Book1")
		This.File = Evl(This.File, plcFile)
		This.File = This.File + Iif(Empty(Justext(This.File)), ".XML", "")

		If Empty(Alias())
			Messagebox("No table is open in the current work area.   ", 48)
			&lcSelectAlias
			Return .F.
		Endif

		*- No Excel casas decimais obrigatóriamente trabalham com ponto "."
		lcSetPoint = Set("Point")
		Set Point To "."
		
		If !This.BuildGridObject(1)
			Return .f.
		Endif
		
		Try	
			lnBytes = 0
			lnBytes = lnBytes + Strtofile("", This.File, 0)
			lnBytes = lnBytes + Strtofile( This.BuildXmlStart() + This.crlf, This.File, 1)
			lnBytes = lnBytes + Strtofile( This.BuildDocumentProperties() + This.crlf, This.File, 1)
			lnBytes = lnBytes + Strtofile( This.BuildXmlExcelWorkbookHeader() + This.crlf, This.File, 1)
			lnBytes = lnBytes + Strtofile( This.BuildAllStyles() + This.crlf, This.File, 1)
			
			Local lnx, lcCaption
			lnx = 1
			lcCaption = allt(This.ProgresBar.Caption)
			For lnx = 1 to This.GridCollection.Count
				If lnx > 1 and !This.BuildGridObject(lnx)
					&lcSelectAlias
					Return .f.
				Endif
				This.ProgresBar.Caption = Textmerge("<<lcCaption>> - (<<This.SheetName>>)")
				
				lnBytes = lnBytes + Strtofile( This.BuildWorksheetHeader() + This.crlf, This.File, 1)

				lnBytes = lnBytes + This.BuildRows()

				lnBytes = lnBytes + Strtofile( This.BuildWorksheetFooter() + This.crlf, This.File, 1)
				
				If Used(This.Alias) AND Reccount(This.Alias) > This._Recno
					Select (This.Alias)
					Go This._Recno
				Endif

			Endfor
			
			wait Window "" TimeOut 1
			This.ProgresBar.Release()
			
			lnBytes = lnBytes + Strtofile( This.BuildXmlExcelWorkbookFooter() + This.crlf, This.File, 1)
			llReturn = Iif(lnBytes > 0, .T., .F.)

		Catch To loError
			Set Step On 
			
			If File(This.File)
				Erase (This.File)
			Endif

			Messagebox("An error occurred during the data exporting. " + Chr(13) + "Error: " + loError.Message, 16, "Exporting")

			llReturn = .F.
		Endtry

		*select (This.StyleCursor)
		*browse normal 
		*select xxxStylesProperties
		*browse normal 

		Set Point To &lcSetPoint

		If Used("xxxStylesProperties")
			Use In xxxStylesProperties
		Endif
		
		Local lnx
		lnx = 1
		For lnx = 1 to This.GridCollection.count
			
			If !This.BuildGridObject(lnx)
				&lcSelectAlias
				Return .f.
			EndIf
			
			If Used(This.StyleCursor)
				Use In (This.StyleCursor)
			EndIf
			
			If Used(This.Alias) AND Reccount(This.Alias) > This._Recno
				Select (This.Alias)
				Go This._Recno
			Endif
		Endfor

		If Vartype(This.GridObject) <> "O"
			loForm.Release()
		Endif

		&lcSelectAlias

		This.GridObject = .Null.
		This.GridCollection = .Null.
				
		*- abre o arquivo apos salva-lo
		If llReturn And This.OpenAfterSaving
			Declare Integer ShellExecute In SHELL32.Dll As WinAPI_OpenAfterSavingExcelXml;
				Integer HndWin, String cAction, String cFileName, ;
				String cParams, String cDir, Integer nShowWin

			WinAPI_OpenAfterSavingExcelXml(0, "OPEN", This.File, "", "", 1)
			Clear Dlls "WinAPI_OpenAfterSavingExcelXml"
		Endif

		Return llReturn

	Endproc

	*|================================================================================ 
	*| ExcelXml::
	Procedure SeekStyle(plcRow, plcCol)

		Local lcReturn, lcStyleCursor
		lcReturn = ""
		lcStyleCursor = This.StyleCursor
		
		*- se nao aplica estilos
		If Not This.SetStyles And plcRow > "000001"
			plcRow = "000001"
		Endif

		*- Procuro um estilo para a celula, caso nao encontre aplico o padrï¿½o.
		*- Teoricamente todas as celulas deve ter um estilo e nï¿½o o padrï¿½o.		 
		If Seek(plcRow + plcCol, lcStyleCursor, "idxRowCol")
			lcReturn = &lcStyleCursor..ssCode
		Else
			lcReturn = "Default"
		Endif

		Return lcReturn

	EndProc
	
	
	*---------------------------------------------------------------------------------------
	* After creating XML file in the Save() method, you can call this method and pass filename of XML file,
	* to use Excel to open the XML file and convert it to an XLSX file.
	*  Values for lnFileFormat:
	* 	51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
	* 	52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
	* 	50 = xlExcel12 (Excel Binary Workbook in 2007-2013 with or without macro's, xlsb)
	* 	56 = xlExcel8 (97-2003 format in Excel 2007-2013, xls)
	Procedure ConvertXmlToXlsx(tcFilename, tnFileFormat, tlOpenAfterExporting)

		Local loExcel as "Excel.Application"
		Local lcNewFilename, lnFileFormat, loWorkBook, lcSafety

		loExcel = Createobject("Excel.Application")

		If Type("loExcel") != "O"
			This.cErrorMessage = "Error starting Excel."
			Return .F.
		Endif

		If !File(tcFileName)
			This.cErrorMessage = "File not found: " + tcFilename
			Return .F.
		Else
			loWorkBook = loExcel.Application.Workbooks.Open(tcFileName)
		EndIf

		lnFileFormat = Evl(tnFileFormat, 51) && 51 = xlsx as default
		
		If (".XML" $ Upper(tcFilename))
			lcNewFilename = Strtran(tcFilename, ".xml", ".xlsx", 1, 99, 1)
			loWorkBook.SaveAs(lcNewFilename, lnFileFormat)
			lcSafety = Set("Safety")
			Set Safety Off
			Delete File (tcFileName)
			Set Safety &lcSafety
		Endif

		If tlOpenAfterExporting
			loExcel.Visible = .T.
		Else
			loExcel.Quit()
		EndIf
		
	Endproc	
	
	*---------------------------------------------------------------------------------------
	* (DCA) - 06/01/2024
	* After creating XML file in the Save() method, you can call this method and pass filename of XML file,
	* to use Excel to open the XML file and convert it to an Any file Format.
	*  Values for lnFileFormat:
	* 	51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
	* 	52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
	* 	50 = xlExcel12 (Excel Binary Workbook in 2007-2013 with or without macro's, xlsb)
	* 	56 = xlExcel8 (97-2003 format in Excel 2007-2013, xls)
	Procedure ConvertXmlToAny(tcFilename, tcFileFormat, tlOverWrite, tlOpenAfterExporting)

		Local loExcel as "Excel.Application"
		Local lcNewFilename, loFileFormat, lcFileFormat, lnFileFormat, loWorkBook, lcSafety, llOverWrite, llFileConverted
		
		loExcel = This.Excel

		llFileConverted = .f.
	
		If Type("loExcel") != "O"
			This.cErrorMessage = "Error starting Excel."
			Return .F.
		Endif
		
		If !File(tcFileName)
			This.cErrorMessage = "File not found: " + tcFilename
			Return .F.
		Else
			loWorkBook = loExcel.Application.Workbooks.Open(tcFileName)
		EndIf

		lcFileFormat = Evl(tcFileFormat,".xlsx")					&& .xlsx 	= xlsx as default
		
		If "xls" == Alltrim(Lower(Strtran(lcFileFormat,".","")))
			loFileFormat = this.GetXlFileFormatEnum(lcFileFormat,5) && Pegar a 5ª iteração
		Else
			loFileFormat = this.GetXlFileFormatEnum(lcFileFormat,1)	&& Pegar a 1ª iteração
		Endif
		
		lnFileFormat = loFileFormat.Value
		
		If (".XML" $ Upper(tcFilename))
			lcNewFilename =  ForceExt(tcFilename, lcFileFormat) &&Strtran(tcFilename, ".xml", lcFileFormat, 1, 99, 1)
			
			If File(lcNewFilename) and tlOverWrite
				Delete File (lcNewFilename)	
			Endif
			
			lcSafety = Set("Safety")
			Set Safety Off
			
			llOverWrite = File(lcNewFilename) AND Pcount()<3 AND 6 = MessageBox(Textmerge("Já existe o arquivo <<lcNewFilename>>.<<Chr(13)+Chr(10)>> Deseja sobrescrever o arquivo?"), 4+32+256, "Confirmation at "+Program())
			
			If llOverWrite
				Delete File (lcNewFilename)	
			EndIf
			
			
			If not File(lcNewFilename)
				loWorkBook.SaveAs(	lcNewFilename	,; && FileName 
									lnFileFormat	,; && FileFormat
									.null.			,; && Password
									.null.			,; && WriteResPassword
									.null.			,; && ReadOnlyRecommended
									.null.			,; && CreateBackup
									.null.			,; && AccessMode
									.null.			,; && ConflictResolution
									.null.			,; && AddToMru
									.null.			,; && TextCodepage
									.null. 			,; && TextVisualLayout
									.t.				)  && Local
				
				llFileConverted = .t.
			Else
				loWorkBook.Close(.f.)
				tlOpenAfterExporting = .f.
			EndIf
						
			Delete File (tcFileName)
			Set Safety &lcSafety
		Endif

		loExcel.Quit()	
		
		*- abre o arquivo apos salva-lo
		If tlOpenAfterExporting
			Declare Integer ShellExecute In SHELL32.Dll As WinAPI_OpenAfterSavingExcelXml;
				Integer HndWin, String cAction, String cFileName, ;
				String cParams, String cDir, Integer nShowWin

			WinAPI_OpenAfterSavingExcelXml(0, "OPEN", lcNewFilename, "", "", 1)
			Clear Dlls "WinAPI_OpenAfterSavingExcelXml"
		EndIf
		
		Release loWorkBook, loExcel
		
		Return llFileConverted
	Endproc	
	
	
	*/ (DCA) - 06/01/2024 */
	*** <summary>
	*** Devolve o XlFileFormat enumeration (Excel) 
	*** </summary>
	*** <remarks></remarks>
	Procedure GetXlFileFormatEnum(tcFileFormat, tnMatch)
		
		Local laXlFileFrmEnum(56) , laFileExten(1), lcFileFormat, lnMatch, lnTempMatch, loXlFileFrmEnum, x, y
		
		lcFileFormat = Alltrim(Evl(tcFileFormat,".xlsx"))
		lcFileFormat = Iif(Left(lcFileFormat,1) <> ".", "."+lcFileFormat, lcFileFormat)
		lnMatch 	 = Evl(Int(Val(Transform(tnMatch))),1) 
		lnTempMatch  = 1
		
		laXlFileFrmEnum[01] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[01],"Name"			,"xlAddIn")
		AddProperty(laXlFileFrmEnum[01],"Value"			,18)
		AddProperty(laXlFileFrmEnum[01],"Description"	,"Microsoft Excel 97-2003 Add-In")
		AddProperty(laXlFileFrmEnum[01],"Extension"		,".xla")
	
		laXlFileFrmEnum[02] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[02],"Name"			,"xlAddIn8")
		AddProperty(laXlFileFrmEnum[02],"Value"			,18)
		AddProperty(laXlFileFrmEnum[02],"Description"	,"Microsoft Excel 97-2003 Add-In")
		AddProperty(laXlFileFrmEnum[02],"Extension"		,".xla")
		
		laXlFileFrmEnum[03] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[03],"Name"			,"xlCSV")
		AddProperty(laXlFileFrmEnum[03],"Value"			,6)
		AddProperty(laXlFileFrmEnum[03],"Description"	,"CSV")
		AddProperty(laXlFileFrmEnum[03],"Extension"		,".csv")
		
		laXlFileFrmEnum[04] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[04],"Name"			,"xlCSVMac")
		AddProperty(laXlFileFrmEnum[04],"Value"			,22)
		AddProperty(laXlFileFrmEnum[04],"Description"	,"Macintosh CSV")
		AddProperty(laXlFileFrmEnum[04],"Extension"		,".csv")
		
		laXlFileFrmEnum[05] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[05],"Name"			,"xlCSVMSDOS")
		AddProperty(laXlFileFrmEnum[05],"Value"			,24)
		AddProperty(laXlFileFrmEnum[05],"Description"	,"MSDOS CSV")
		AddProperty(laXlFileFrmEnum[05],"Extension"		,".csv")
		
		laXlFileFrmEnum[06] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[06],"Name"			,"xlCSVUTF8")
		AddProperty(laXlFileFrmEnum[06],"Value"			,62)
		AddProperty(laXlFileFrmEnum[06],"Description"	,"UTF8 CSV")
		AddProperty(laXlFileFrmEnum[06],"Extension"		,".csv")
		
		laXlFileFrmEnum[07] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[07],"Name"			,"xlCSVWindows")
		AddProperty(laXlFileFrmEnum[07],"Value"			,23)
		AddProperty(laXlFileFrmEnum[07],"Description"	,"Windows CSV")
		AddProperty(laXlFileFrmEnum[07],"Extension"		,".csv")
		
		laXlFileFrmEnum[08] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[08],"Name"			,"xlCurrentPlatformText")
		AddProperty(laXlFileFrmEnum[08],"Value"			,-4158)
		AddProperty(laXlFileFrmEnum[08],"Description"	,"Current Platform Text")
		AddProperty(laXlFileFrmEnum[08],"Extension"		,".txt")
		
		laXlFileFrmEnum[09] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[09],"Name"			,"xlDBF2")
		AddProperty(laXlFileFrmEnum[09],"Value"			,7)
		AddProperty(laXlFileFrmEnum[09],"Description"	,"Dbase 2 format")
		AddProperty(laXlFileFrmEnum[09],"Extension"		,".dbf")
		
		laXlFileFrmEnum[10] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[10],"Name"			,"xlDBF3")
		AddProperty(laXlFileFrmEnum[10],"Value"			,8)
		AddProperty(laXlFileFrmEnum[10],"Description"	,"Dbase 3 format")
		AddProperty(laXlFileFrmEnum[10],"Extension"		,".dbf")
		
		laXlFileFrmEnum[11] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[11],"Name"			,"xlDBF4")
		AddProperty(laXlFileFrmEnum[11],"Value"			,11)
		AddProperty(laXlFileFrmEnum[11],"Description"	,"Dbase 4 format")
		AddProperty(laXlFileFrmEnum[11],"Extension"		,".dbf")
		
		laXlFileFrmEnum[12] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[12],"Name"			,"xlDIF")
		AddProperty(laXlFileFrmEnum[12],"Value"			,9)
		AddProperty(laXlFileFrmEnum[12],"Description"	,"Data Interchange format")
		AddProperty(laXlFileFrmEnum[12],"Extension"		,".dif")
		
		laXlFileFrmEnum[13] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[13],"Name"			,"xlExcel12")
		AddProperty(laXlFileFrmEnum[13],"Value"			,50)
		AddProperty(laXlFileFrmEnum[13],"Description"	,"Excel Binary Workbook")
		AddProperty(laXlFileFrmEnum[13],"Extension"		,".xlsb")
		
		laXlFileFrmEnum[14] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[14],"Name"			,"xlExcel2")
		AddProperty(laXlFileFrmEnum[14],"Value"			,16)
		AddProperty(laXlFileFrmEnum[14],"Description"	,"Excel version 2.0 (1987)")
		AddProperty(laXlFileFrmEnum[14],"Extension"		,".xls")
		
		laXlFileFrmEnum[15] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[15],"Name"			,"xlExcel2FarEast")
		AddProperty(laXlFileFrmEnum[15],"Value"			,27)
		AddProperty(laXlFileFrmEnum[15],"Description"	,"Excel version 2.0 Asia (1987)")
		AddProperty(laXlFileFrmEnum[15],"Extension"		,".xls")
		
		laXlFileFrmEnum[16] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[16],"Name"			,"xlExcel3")
		AddProperty(laXlFileFrmEnum[16],"Value"			,29)
		AddProperty(laXlFileFrmEnum[16],"Description"	,"Excel version 3.0 (1990)")
		AddProperty(laXlFileFrmEnum[16],"Extension"		,".xls")
		
		laXlFileFrmEnum[17] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[17],"Name"			,"xlExcel4")
		AddProperty(laXlFileFrmEnum[17],"Value"			,33)
		AddProperty(laXlFileFrmEnum[17],"Description"	,"Excel version 4.0 (1992)")
		AddProperty(laXlFileFrmEnum[17],"Extension"		,".xls")
		
		laXlFileFrmEnum[18] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[18],"Name"			,"xlExcel4Workbook")
		AddProperty(laXlFileFrmEnum[18],"Value"			,35)
		AddProperty(laXlFileFrmEnum[18],"Description"	,"Excel version 4.0. Workbook format (1992)")
		AddProperty(laXlFileFrmEnum[18],"Extension"		,".xlw")
		
		laXlFileFrmEnum[19] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[19],"Name"			,"xlExcel5")
		AddProperty(laXlFileFrmEnum[19],"Value"			,39)
		AddProperty(laXlFileFrmEnum[19],"Description"	,"Excel version 5.0 (1994)")
		AddProperty(laXlFileFrmEnum[19],"Extension"		,".xls")
		
		laXlFileFrmEnum[20] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[20],"Name"			,"xlExcel7")
		AddProperty(laXlFileFrmEnum[20],"Value"			,39)
		AddProperty(laXlFileFrmEnum[20],"Description"	,"Excel 95 (version 7.0)")
		AddProperty(laXlFileFrmEnum[20],"Extension"		,".xls")
		
		laXlFileFrmEnum[21] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[21],"Name"			,"xlExcel8")
		AddProperty(laXlFileFrmEnum[21],"Value"			,56)
		AddProperty(laXlFileFrmEnum[21],"Description"	,"Excel 97-2003 Workbook")
		AddProperty(laXlFileFrmEnum[21],"Extension"		,".xls")
		
		laXlFileFrmEnum[22] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[22],"Name"			,"xlExcel9795")
		AddProperty(laXlFileFrmEnum[22],"Value"			,43)
		AddProperty(laXlFileFrmEnum[22],"Description"	,"Excel version 95 and 97")
		AddProperty(laXlFileFrmEnum[22],"Extension"		,".xls")
		
		laXlFileFrmEnum[23] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[23],"Name"			,"xlHtml")
		AddProperty(laXlFileFrmEnum[23],"Value"			,44)
		AddProperty(laXlFileFrmEnum[23],"Description"	,"HTML format")
		AddProperty(laXlFileFrmEnum[23],"Extension"		,".htm;.html")
		
		laXlFileFrmEnum[24] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[24],"Name"			,"xlIntlAddIn")
		AddProperty(laXlFileFrmEnum[24],"Value"			,26)
		AddProperty(laXlFileFrmEnum[24],"Description"	,"International Add-In")
		AddProperty(laXlFileFrmEnum[24],"Extension"		,"No file extension")
		
		laXlFileFrmEnum[25] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[25],"Name"			,"xlIntlMacro")
		AddProperty(laXlFileFrmEnum[25],"Value"			,25)
		AddProperty(laXlFileFrmEnum[25],"Description"	,"International Macro")
		AddProperty(laXlFileFrmEnum[25],"Extension"		,"No file extension")
		
		laXlFileFrmEnum[26] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[26],"Name"			,"xlOpenDocumentSpreadsheet")
		AddProperty(laXlFileFrmEnum[26],"Value"			,60)
		AddProperty(laXlFileFrmEnum[26],"Description"	,"OpenDocument Spreadsheet")
		AddProperty(laXlFileFrmEnum[26],"Extension"		,".ods")
		
		laXlFileFrmEnum[27] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[27],"Name"			,"xlOpenXMLAddIn")
		AddProperty(laXlFileFrmEnum[27],"Value"			,55)
		AddProperty(laXlFileFrmEnum[27],"Description"	,"Open XML Add-In")
		AddProperty(laXlFileFrmEnum[27],"Extension"		,".xlam")
		
		laXlFileFrmEnum[28] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[28],"Name"			,"xlOpenXMLStrictWorkbook")
		AddProperty(laXlFileFrmEnum[28],"Value"			,61) 
		AddProperty(laXlFileFrmEnum[28],"Description"	,"Strict Open XML file")
		AddProperty(laXlFileFrmEnum[28],"Extension"		,".xlsx")
		
		laXlFileFrmEnum[29] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[29],"Name"			,"xlOpenXMLTemplate")
		AddProperty(laXlFileFrmEnum[29],"Value"			,54) 
		AddProperty(laXlFileFrmEnum[29],"Description"	,"Open XML Template")
		AddProperty(laXlFileFrmEnum[29],"Extension"		,".xltx")
		
		laXlFileFrmEnum[30] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[30],"Name"			,"xlOpenXMLTemplateMacroEnabled")
		AddProperty(laXlFileFrmEnum[30],"Value"			,53) 
		AddProperty(laXlFileFrmEnum[30],"Description"	,"Open XML Template Macro Enabled")
		AddProperty(laXlFileFrmEnum[30],"Extension"		,".xltm")
		
		laXlFileFrmEnum[31] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[31],"Name"			,"xlOpenXMLWorkbook")
		AddProperty(laXlFileFrmEnum[31],"Value"			,51) 
		AddProperty(laXlFileFrmEnum[31],"Description"	,"Open XML Workbook")
		AddProperty(laXlFileFrmEnum[31],"Extension"		,".xlsx")
		
		laXlFileFrmEnum[32] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[32],"Name"			,"xlOpenXMLWorkbookMacroEnabled")
		AddProperty(laXlFileFrmEnum[32],"Value"			,52) 
		AddProperty(laXlFileFrmEnum[32],"Description"	,"Open XML Workbook Macro Enabled")
		AddProperty(laXlFileFrmEnum[32],"Extension"		,".xlsm")
		
		laXlFileFrmEnum[33] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[33],"Name"			,"xlSYLK")
		AddProperty(laXlFileFrmEnum[33],"Value"			,2) 
		AddProperty(laXlFileFrmEnum[33],"Description"	,"Symbolic Link format")
		AddProperty(laXlFileFrmEnum[33],"Extension"		,".slk")
		
		laXlFileFrmEnum[34] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[34],"Name"			,"xlTemplate")
		AddProperty(laXlFileFrmEnum[34],"Value"			,17) 
		AddProperty(laXlFileFrmEnum[34],"Description"	,"Excel Template format")
		AddProperty(laXlFileFrmEnum[34],"Extension"		,".xlt")
		
		laXlFileFrmEnum[35] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[35],"Name"			,"xlTemplate8")
		AddProperty(laXlFileFrmEnum[35],"Value"			,17) 
		AddProperty(laXlFileFrmEnum[35],"Description"	,"Template 8")
		AddProperty(laXlFileFrmEnum[35],"Extension"		,".xlt")
		
		laXlFileFrmEnum[36] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[36],"Name"			,"xlTextMac")
		AddProperty(laXlFileFrmEnum[36],"Value"			,19) 
		AddProperty(laXlFileFrmEnum[36],"Description"	,"Macintosh Text")
		AddProperty(laXlFileFrmEnum[36],"Extension"		,".txt")
		
		laXlFileFrmEnum[37] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[37],"Name"			,"xlTextMSDOS")
		AddProperty(laXlFileFrmEnum[37],"Value"			,21) 
		AddProperty(laXlFileFrmEnum[37],"Description"	,"MSDOS Text")
		AddProperty(laXlFileFrmEnum[37],"Extension"		,".txt")
		
		laXlFileFrmEnum[38] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[38],"Name"			,"xlTextPrinter")
		AddProperty(laXlFileFrmEnum[38],"Value"			,36) 
		AddProperty(laXlFileFrmEnum[38],"Description"	,"Printer Text")
		AddProperty(laXlFileFrmEnum[38],"Extension"		,".prn")
		
		laXlFileFrmEnum[39] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[39],"Name"			,"xlTextWindows")
		AddProperty(laXlFileFrmEnum[39],"Value"			,20) 
		AddProperty(laXlFileFrmEnum[39],"Description"	,"Windows Text")
		AddProperty(laXlFileFrmEnum[39],"Extension"		,".txt")
		
		laXlFileFrmEnum[40] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[40],"Name"			,"xlUnicodeText")
		AddProperty(laXlFileFrmEnum[40],"Value"			,42) 
		AddProperty(laXlFileFrmEnum[40],"Description"	,"Unicode Text")
		AddProperty(laXlFileFrmEnum[40],"Extension"		,"No file extension;.txt")
		
		laXlFileFrmEnum[41] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[41],"Name"			,"xlWebArchive")
		AddProperty(laXlFileFrmEnum[41],"Value"			,45) 
		AddProperty(laXlFileFrmEnum[41],"Description"	,"Web Archive")
		AddProperty(laXlFileFrmEnum[41],"Extension"		,".mht;.mhtml")
		
		laXlFileFrmEnum[42] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[42],"Name"			,"xlWJ2WD1")
		AddProperty(laXlFileFrmEnum[42],"Value"			,14) 
		AddProperty(laXlFileFrmEnum[42],"Description"	,"Japanese 1-2-3")
		AddProperty(laXlFileFrmEnum[42],"Extension"		,".wj2")
		
		laXlFileFrmEnum[43] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[43],"Name"			,"xlWJ3")
		AddProperty(laXlFileFrmEnum[43],"Value"			,40) 
		AddProperty(laXlFileFrmEnum[43],"Description"	,"Japanese 1-2-3")
		AddProperty(laXlFileFrmEnum[43],"Extension"		,".wj3")
		
		laXlFileFrmEnum[44] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[44],"Name"			,"xlWJ3FJ3")
		AddProperty(laXlFileFrmEnum[44],"Value"			,41) 
		AddProperty(laXlFileFrmEnum[44],"Description"	,"Japanese 1-2-3 format")
		AddProperty(laXlFileFrmEnum[44],"Extension"		,".wj3")
		
		laXlFileFrmEnum[45] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[45],"Name"			,"xlWK1")
		AddProperty(laXlFileFrmEnum[45],"Value"			,5) 
		AddProperty(laXlFileFrmEnum[45],"Description"	,"Lotus 1-2-3 format")
		AddProperty(laXlFileFrmEnum[45],"Extension"		,".wk1")
		
		laXlFileFrmEnum[46] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[46],"Name"			,"xlWK1ALL")
		AddProperty(laXlFileFrmEnum[46],"Value"			,31) 
		AddProperty(laXlFileFrmEnum[46],"Description"	,"Lotus 1-2-3 format")
		AddProperty(laXlFileFrmEnum[46],"Extension"		,".wk1")
		
		laXlFileFrmEnum[47] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[47],"Name"			,"xlWK1FMT")
		AddProperty(laXlFileFrmEnum[47],"Value"			,30) 
		AddProperty(laXlFileFrmEnum[47],"Description"	,"Lotus 1-2-3 format")
		AddProperty(laXlFileFrmEnum[47],"Extension"		,".wk1")
		
		laXlFileFrmEnum[48] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[48],"Name"			,"xlWK3")
		AddProperty(laXlFileFrmEnum[48],"Value"			,15) 
		AddProperty(laXlFileFrmEnum[48],"Description"	,"Lotus 1-2-3 format")
		AddProperty(laXlFileFrmEnum[48],"Extension"		,".wk3")
		
		laXlFileFrmEnum[49] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[49],"Name"			,"xlWK3FM3")
		AddProperty(laXlFileFrmEnum[49],"Value"			,32) 
		AddProperty(laXlFileFrmEnum[49],"Description"	,"Lotus 1-2-3 format")
		AddProperty(laXlFileFrmEnum[49],"Extension"		,".wk3")
		
		laXlFileFrmEnum[50] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[50],"Name"			,"xlWK4")
		AddProperty(laXlFileFrmEnum[50],"Value"			,38) 
		AddProperty(laXlFileFrmEnum[50],"Description"	,"Lotus 1-2-3 format")
		AddProperty(laXlFileFrmEnum[50],"Extension"		,".wk4")
		
		laXlFileFrmEnum[51] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[51],"Name"			,"xlWKS")
		AddProperty(laXlFileFrmEnum[51],"Value"			,4) 
		AddProperty(laXlFileFrmEnum[51],"Description"	,"Lotus 1-2-3 format")
		AddProperty(laXlFileFrmEnum[51],"Extension"		,".wks")
		
		laXlFileFrmEnum[52] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[52],"Name"			,"xlWorkbookDefault")
		AddProperty(laXlFileFrmEnum[52],"Value"			,51) 
		AddProperty(laXlFileFrmEnum[52],"Description"	,"Workbook default")
		AddProperty(laXlFileFrmEnum[52],"Extension"		,".xlsx")
		
		laXlFileFrmEnum[53] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[53],"Name"			,"xlWorkbookNormal")
		AddProperty(laXlFileFrmEnum[53],"Value"			,-4143) 
		AddProperty(laXlFileFrmEnum[53],"Description"	,"Workbook normal")
		AddProperty(laXlFileFrmEnum[53],"Extension"		,".xls")
		
		laXlFileFrmEnum[54] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[54],"Name"			,"xlWorks2FarEast")
		AddProperty(laXlFileFrmEnum[54],"Value"			,28) 
		AddProperty(laXlFileFrmEnum[54],"Description"	,"Microsoft Works 2.0 Asian format")
		AddProperty(laXlFileFrmEnum[54],"Extension"		,".wks")
		
		laXlFileFrmEnum[55] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[55],"Name"			,"xlWQ1")
		AddProperty(laXlFileFrmEnum[55],"Value"			,34) 
		AddProperty(laXlFileFrmEnum[55],"Description"	,"Quattro Pro format")
		AddProperty(laXlFileFrmEnum[55],"Extension"		,".wq1")
		
		laXlFileFrmEnum[56] = CreateObject("empty")
		AddProperty(laXlFileFrmEnum[56],"Name"			,"xlXMLSpreadsheet")
		AddProperty(laXlFileFrmEnum[56],"Value"			,46) 
		AddProperty(laXlFileFrmEnum[56],"Description"	,"XML Spreadsheet")
		AddProperty(laXlFileFrmEnum[56],"Extension"		,".xml")
		
		
		loXlFileFrmEnum = laXlFileFrmEnum[31]
		x = 1
		For x = 1 to Alen(laXlFileFrmEnum,1)
			
			ALines(laFileExten, laXlFileFrmEnum[x].Extension,";")
			
			y = 1
			For y = 1 to Alen(laFileExten,1)
				
				If Lower(lcFileFormat) == Lower(laFileExten[y])
					loXlFileFrmEnum = laXlFileFrmEnum[x]
				EndIf
				
				If Lower(lcFileFormat) == Lower(laFileExten[y]) and lnMatch = lnTempMatch
					Return loXlFileFrmEnum
				Endif
				
				If Lower(lcFileFormat) == Lower(laFileExten[y]) and lnMatch > lnTempMatch 
					lnTempMatch = lnTempMatch + 1	
				EndIf
				
			Endfor
			
		Endfor
		
		Return loXlFileFrmEnum
	Endproc
	
Enddefine

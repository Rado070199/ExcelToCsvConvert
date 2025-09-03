DECLARE 
		@CMD VARCHAR(2000)
	,	@ETCC0_sciezka_do_pliku_excel VARCHAR(1000) = 'G:\WymianaDanych\SAP_WymianaDanych\test\plik.xlsx'
	,	@ETCC1_sciezka_do_pliku_csv VARCHAR(1000) = 'G:\WymianaDanych\SAP_WymianaDanych\test\plik.csv'
	,	@ETCC2_separator_csv VARCHAR(10) = ';'
	,	@ETCC3_indeks_skoroszytu_liczony_od_zero VARCHAR(10) = '1'
		CREATE TABLE #output (line NVARCHAR(4000))
	
SET @CMD = 'G:/WymianaDanych/programy/ExcelToCsvConverter/ExcelToCsvConvert.exe '
         + '"' + @ETCC0_sciezka_do_pliku_excel + '" '
         + '"' + @ETCC1_sciezka_do_pliku_csv + '" '
         + '"' + @ETCC2_separator_csv + '" '
         + @ETCC3_indeks_skoroszytu_liczony_od_zero;
PRINT @CMD

INSERT INTO #output EXEC MASTER..XP_CMDSHELL @CMD

select * from #output

DROP TABLE #output
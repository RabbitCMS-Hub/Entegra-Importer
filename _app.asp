<%'/*
'**********************************************
'      /\      | (_)
'     /  \   __| |_  __ _ _ __  ___
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/
'**********************************************
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************
' Class Structure for RabbitCMS PluginManagers
' 	- Private and Public Variables
' 	- class_register()
' 	- LoadPanel()
' 	- class_initialize()
' 	- class_terminate()
' 	- class_register()
' 	- GET PluginCode()
' 	- GET PluginName()
' 	- GET PluginVersion()
' 	- GET PluginCredits()
' 	- GET AutoLoad()

' 	- YOUR SUB/FUNC/PROPERTY
' 	- YOUR SUB/FUNC/PROPERTY
' 	- YOUR SUB/FUNC/PROPERTY
'**********************************************
' Response.Charset = "UTF-8"
' Response.Codepage = 65001
' Response.codepage = 1254
' Response.charset = "windows-1254"
'*/
Class Entegra_Importer
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Variables
	'---------------------------------------------------------------
	'*/
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Variables
	'---------------------------------------------------------------
	'*/

    Private YesOverWrite, NoOverWrite
    Private DEBUG, SOURCE_URL, SOURCE_CODE, DOWNLOAD_WITH, PROTECT_ORIGINAL_FILE, DOWNLOAD_FILE_PATH

	'/*
	'---------------------------------------------------------------
	' REQUIRED: Register Class
	'---------------------------------------------------------------
	'*/
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		'/*
		'---------------------------------------------------------------
		' Check Register
		'---------------------------------------------------------------
		'*/
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If
		'/*
		'---------------------------------------------------------------
		' Plugin Database
		'---------------------------------------------------------------
		'*/
		' Dim PluginTableName
		' 	PluginTableName = "tbl_plugin_" & PLUGIN_DB_NAME
    	
  '   	If TableExist(PluginTableName) = False Then
		' 	DebugTimer ""& PLUGIN_CODE &" table creating"
    		
  '   		Conn.Execute("SET NAMES utf8mb4;") 
  '   		Conn.Execute("SET FOREIGN_KEY_CHECKS = 0;") 
    		
  '   		Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"`")

  '   		q="CREATE TABLE `"& PluginTableName &"` ( "
  '   		q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
  '   		q=q+"  `FILENAME` varchar(255) DEFAULT NULL, "
  '   		q=q+"  `FULL_PATH` varchar(255) DEFAULT NULL, "
  '   		q=q+"  `COMPRESS_DATE` datetime DEFAULT NULL, "
  '   		q=q+"  `COMPRESS_RATIO` double(255,0) DEFAULT NULL, "
  '   		q=q+"  `ORIGINAL_FILE_SIZE` bigint(20) DEFAULT 0, "
  '   		q=q+"  `COMPRESSED_FILE_SIZE` bigint(20) DEFAULT 0, "
  '   		q=q+"  `EARNED_SIZE` bigint(20) DEFAULT 0, "
  '   		q=q+"  `ORIGINAL_PROTECTED` int(1) DEFAULT 0, "
  '   		q=q+"  PRIMARY KEY (`ID`), "
  '   		q=q+"  KEY `IND1` (`FILENAME`) "
  '   		q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
		' 	Conn.Execute(q)

  '   		Conn.Execute("SET FOREIGN_KEY_CHECKS = 1;") 

		' 	' Create Log
		' 	'------------------------------
  '   		Call PanelLog(""& PLUGIN_CODE &" için database tablosu oluşturuldu", 0, ""& PLUGIN_CODE &"", 0)

		' 	' Register Settings
		' 	'------------------------------
		' 	DebugTimer ""& PLUGIN_CODE &" class_register() End"
  '   	End If
		'/*
		'---------------------------------------------------------------
		' Plugin Settings
		'---------------------------------------------------------------
		'*/
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE&"_")
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "Entegra_Importer")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "0")
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		'/*
		'---------------------------------------------------------------
		' Register Settings
		'---------------------------------------------------------------
		'*/
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Register Class End
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Settings Panel
	'---------------------------------------------------------------
	'*/
	Public sub LoadPanel()
		'/*
		'--------------------------------------------------------
		' Sub Page
		'--------------------------------------------------------
		'*/
		If Query.Data("Page") = "SHOW:CachedFiles" Then
			Call PluginPage("Header")

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If
		'/*
		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		'*/
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			PLUGIN_PANEL_INPUT(This, "select", "OPTION_1", "Buraya Title", "0#Seçenek 1|1#Seçenek 2|2#Seçenek 3", TO_DB)
			' .Write 			QuickSettings("select", ""& PLUGIN_CODE &"_OPTION_1", "Buraya Title", "0#Seçenek 1|1#Seçenek 2|2#Seçenek 3", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("number", ""& PLUGIN_CODE &"_OPTION_2", "Buraya Title", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write 			QuickSettings("tag", ""& PLUGIN_CODE &"_OPTION_3", "Buraya Title", "", TO_DB)
			.Write "    </div>"
			.Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:CachedFiles"" class=""btn btn-sm btn-primary"">"
			.Write "        	Önbelleklenmiş Dosyaları Göster"
			.Write "        </a>"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=DELETE:CachedFiles"" class=""btn btn-sm btn-danger"">"
			.Write "        	Tüm Önbelleği Temizle"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Settings Panel
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Initialize
	'---------------------------------------------------------------
	'*/
	Private Sub class_initialize()
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: PluginTemplate Main Variables
		'-----------------------------------------------------------------------------------
		'*/
    	PLUGIN_CODE  			= "ENTEGRA_IMPORTER"
    	PLUGIN_NAME 			= "Entegra Importer"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Entegra-Importer"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-archive"
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_FOLDER_NAME 		= "Entegra-Importer"
    	PLUGIN_DB_NAME 			= "plugin_entegra_importer"
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_AUTOLOAD 		= True
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
		'/*
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
		'*/

        Response.Buffer         = true
        Server.ScriptTimeout    = 60 * 60

        YesOverWrite            = 2
        NoOverWrite             = 1
        DEBUG                   = False ' Verbose to Screen
        
        SOURCE_URL              = Null
        SOURCE_CODE             = Null
        DOWNLOAD_WITH           = "ASPJPEG" '"ADODB"
        PROTECT_ORIGINAL_FILE   = False
        
        DOWNLOAD_FILE_PATH      = Server.MapPath( CMS_PRODUCT_FILES_ROOT & Year(Now()) &"/"& Month(Now()) ) &"\"

		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: Register Plugin to CMS
		'-----------------------------------------------------------------------------------
		'*/
		class_register()
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: Hook Plugin to CMS Auto Load Location WEB|API|PANEL
		'-----------------------------------------------------------------------------------
		'*/
		If PLUGIN_AUTOLOAD_AT("WEB") = True Then 
			' Cms.BodyData = Init()
			' Cms.FooterData = "<add-footer-html>Hello World!</add-footer-html>"
		End If
	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Initialize
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Terminate
	'---------------------------------------------------------------
	'*/
	Private sub class_terminate()

	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Terminate
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Manager Exports
	'---------------------------------------------------------------
	'*/
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property
	Public Property Get PluginAutoload() 	: PluginAutoload = PLUGIN_AUTOLOAD 			: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable, PluginAutoload)
	End Property
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Manager Exports
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Function LoadXmlData(data)
        Set LoadXmlData = CreateObject("MSXML2.FreeThreadedDOMDocument")
            LoadXmlData.setProperty "SelectionLanguage", "XPath"
            LoadXmlData.LoadXML(data)
            If LoadXmlData.parseError <> 0 Then
                Err.Raise vbObjectError + 1, _
                    "LoadXmlDocument", _
                    "Cannot load " & url & " (" & LoadXmlData.parseError.reason & ")"
            End If
    End Function
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Function Fetch()
        Set objHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0") 
        With objHTTP
            .Open "GET", SOURCE_URL, False
            .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
            .setTimeouts 50000, 50000, 50000, 50000
            .send
        End With

        If objHTTP.Status = 200 Then 
            Verbose "XML", "Veriler Alındı."
            SOURCE_CODE = objHTTP.responseText
        Else 
            Verbose "XML", "Veriler Alınamadı."
            SOURCE_CODE = Null
        End If
        Set objHTTP = Nothing
    End Function
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Function GetTextHTMLEncode(context, xpath)
        Dim node
        Set node = context.selectSingleNode(xpath)
        If node Is Nothing Then
            GetTextHTMLEncode = Null
        Else
            Dim tmp_data
                tmp_data = node.text
                tmp_data = Server.HTMLEncode( tmp_data  )
                tmp_data = TurkceKarakterRevert( tmp_data )
            GetTextHTMLEncode = LoginKontrol(Trim(tmp_data))
        End If
        Set node = Nothing
    End Function
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Function TurkceKarakterRevert(Text) 
        Gelen2 = Text  
        Gelen2 = Replace(Gelen2,"&#286;" ,"Ğ" )  
        Gelen2 = Replace(Gelen2,"&#287;" ,"ğ" )  
        Gelen2 = Replace(Gelen2,"&#214;" ,"Ö" )  
        Gelen2 = Replace(Gelen2,"&#246;" ,"ö" )  
        Gelen2 = Replace(Gelen2,"&#304;" ,"İ" )  
        Gelen2 = Replace(Gelen2,"&#305;" ,"ı" )  
        Gelen2 = Replace(Gelen2,"&#220;" ,"Ü" )  
        Gelen2 = Replace(Gelen2,"&#252;" ,"ü" )  
        Gelen2 = Replace(Gelen2,"&#350;" ,"Ş" )  
        Gelen2 = Replace(Gelen2,"&#351;" ,"ş" )  
        Gelen2 = Replace(Gelen2,"&#199;" ,"Ç" )  
        Gelen2 = Replace(Gelen2,"&#231;" ,"ç" )
        
        Gelen2 = Replace(Gelen2,"&#62;" ,">" )  
        Gelen2 = Replace(Gelen2,"&#60;" ,"<" ) 
        Gelen2 = Replace(Gelen2,"&#32;" ," " ) 
        Gelen2 = Replace(Gelen2,"&#33;" ,"!" ) 
        Gelen2 = Replace(Gelen2,"&#34;" ,Chr(34) ) 
        Gelen2 = Replace(Gelen2,"&#35;" ,"#" ) 
        Gelen2 = Replace(Gelen2,"&#36;" ,"$" ) 
        'Gelen2 = Replace(Gelen2,"&#37;" ,"%" ) 
        Gelen2 = Replace(Gelen2,"&#38;" ,"&" ) 
        Gelen2 = Replace(Gelen2,"&#39;" ,"'" ) 
        Gelen2 = Replace(Gelen2,"&#61;" ,"=" ) 
        Gelen2 = Replace(Gelen2,"&#63;" ,"?" ) 
        Gelen2 = Replace(Gelen2,"&#64;" ,"@" ) 
        Gelen2 = Replace(Gelen2,"&#91;" ,"[" ) 
        Gelen2 = Replace(Gelen2,"&#92;" ,"\" ) 
        Gelen2 = Replace(Gelen2,"&#93;" ,"]" ) 
        Gelen2 = Replace(Gelen2,"&#94;" ,"^" ) 
        Gelen2 = Replace(Gelen2,"&#95;" ,"_" ) 
        Gelen2 = Replace(Gelen2,"&#96;" ,"`" ) 
        Gelen2 = Replace(Gelen2,"&#123;" ,"{" ) 
        Gelen2 = Replace(Gelen2,"&#124;" ,"|" ) 
        Gelen2 = Replace(Gelen2,"&#125;" ,"}" ) 
        Gelen2 = Replace(Gelen2,"&#126;" ,"~" ) 

        ' Invisible Chars
        Gelen2 = Replace(Gelen2,"&#8203;" ,"" ) 

        TurkceKarakterRevert = EntityConvert( Gelen2 )
    End Function
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Function EntityConvert(strText)
        yazi2 = strText
        If yazi2 = "" Then Exit Function 
        
        yazi2 = Replace(yazi2, "&lt;", "<")
        yazi2 = Replace(yazi2, "&gt;", ">")
        yazi2 = Replace(yazi2, "&amp;", "&")
        yazi2 = Replace(yazi2, "&quot;", Chr(34))
        
        yazi2 = Replace(yazi2, "<br>", vbcrlf)
        yazi2 = Replace(yazi2, "<br/>", vbcrlf)
        yazi2 = Replace(yazi2, "<br />", vbcrlf)
        EntityConvert = yazi2
    End Function
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Let SetURL(v)
        SOURCE_URL = v 
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get Source()
        Source = SOURCE_CODE
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/


	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get SetCategory(ProductId, CategoryName)
        Set rsCheck = Conn.Execute("SELECT ID FROM tbl_urun_kategorileri WHERE KATEGORI='"& CategoryName &"'")
        If rsCheck.Eof Then 
            Verbose "CATEGORY", "Kategori Oluşturuldu."
            Conn.Execute("INSERT INTO tbl_urun_kategorileri(KATEGORI, AKTIF) VALUES('"& CategoryName &"', 1)")
            
            SetCategory = Clng( Conn.Execute("SELECT LAST_INSERT_ID()")(0) )
            
            Conn.Execute("INSERT INTO tbl_urun_kategorileri_diller(KATEGORIID, DILID, KATEGORIADI) VALUES('"& SetCategory &"', 1, '"& CategoryName &"')")

            Conn.Execute("INSERT INTO tbl_urun_kategori_secimleri(KATEGORIID, URUNID) VALUES('"& SetCategory &"', '"& ProductId &"')")
        Else 
            Verbose "CATEGORY", "Kategori Mevcut."
            SetCategory = rsCheck("ID").Value
            
            Conn.Execute("DELETE FROM tbl_urun_kategori_secimleri WHERE URUNID="& ProductId &"")
            Conn.Execute("INSERT INTO tbl_urun_kategori_secimleri(KATEGORIID, URUNID) VALUES('"& SetCategory &"', '"& ProductId &"')")
        End If
        rsCheck.Close : Set rsCheck = Nothing
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/


	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get SetProduct(ProductCode, ProductName)
        Set rsCheck = Conn.Execute("SELECT ID FROM tbl_urun WHERE STOKKODU='"& ProductCode &"'")
        If rsCheck.Eof Then 
            Verbose "CREATE", "Ürün Oluşturuldu. "& ProductCode &""
            Conn.Execute("INSERT INTO tbl_urun(URUNADI, STOKKODU, DURUM) VALUES('"& ProductName &"', '"& ProductCode &"', 1)")
            
            SetProduct = Clng( Conn.Execute("SELECT LAST_INSERT_ID()")(0) )
        Else 
            Verbose "UPDATE", "Ürün Güncellendi. "& ProductCode &""
            Conn.Execute("UPDATE tbl_urun SET GUNCELLENME_TARIHI=NOW(), DURUM=1, URUNADI='"& ProductName &"' WHERE ID="& rsCheck("ID") &"")
            
            SetProduct = rsCheck("ID").Value
        End If
        rsCheck.Close : Set rsCheck = Nothing
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get UpdateProduct(ProductId, ScriptingObj)
        Dim SQL_STRING
            SQL_STRING = BuildSQL(ScriptingObj, ProductId)

        If SQL_STRING = False Then 
            Verbose "UPDATE", "SQL Oluşturulamadı."
            UpdateProduct = False
            Exit Property
        End If

        Conn.Execute( SQL_STRING )
        Conn.Execute("UPDATE tbl_urun_varyasyonlar SET FIYAT='"& ScriptingObj.item("FIYAT") &"' WHERE URUN_ID="& ProductId &"")

        Verbose "UPDATE", "Bilgiler Güncellendi."
        
        UpdateProduct = True
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/


	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get UpdatePhoto(ProductId, PhotoURL)
        Dim FotografURL, FileName
            FotografURL = "" & PhotoURL
            FileName    = GetFileName(FotografURL)

        Select Case Left(LCase(FotografURL), 7) ' http://
            Case "http://", "https:/"
                Conn.Execute("DELETE FROM tbl_urun_fotograflar WHERE URUNID="& ProductId &"")

                DOWNLOAD_STATUS = DownloadFile(FotografURL, ProductId)

                If DOWNLOAD_STATUS = True Then 
                    Verbose "FOTO", "İndirildi."

                    FOTOGRAF = SetPhotos(DOWNLOAD_FILE_PATH, FileName)

                    If Len(FOTOGRAF) > 3 Then
                        Conn.Execute("DELETE FROM tbl_urun_fotograflar WHERE URUNID="& ProductId &"")
                        Conn.Execute("INSERT INTO tbl_urun_fotograflar(FOTOGRAF, URUNID, MAINFOTO, YIL, AY, EKLENME_TARIHI) VALUES('"& FOTOGRAF &"', '"& ProductId &"', 1, '"& Year(Now()) &"', '"& Month(Now()) &"', NOW())")
                    End If
                Else 
                    Verbose "FOTO", "İndirilemedi. Kaynak: "& FotografURL &""
                End If
            Case Else 

        End Select
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get Dictionary()
        Set Dictionary = CreateObject("Scripting.Dictionary")
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get DownloadFile(FromURL, ProductId)
        Call ContentKlasoruOlustur

        Dim FileName
            FileName = GetFileName(FromURL)
        
        If FileName = False Then 
            Verbose "DOWNLOAD", "Dosya Adı Bulunamadı, İptal Edildi..."
            DownloadFile = False
            Exit Property
        End If

        If IsFileExist(DOWNLOAD_FILE_PATH & FileName) = True Then 
            ' Set DosyaAc = Conn.Execute("SELECT ID,YIL,AY,FOTOGRAF FROM tbl_urun_fotograflar WHERE URUNID = "& ProductId &"")
            ' Do while Not DosyaAc.Eof
            '     str_file_path = CMS_PRODUCT_FILES_ROOT & ""&DosyaAc("YIL")&"/"&DosyaAc("AY")&"/"
            '     On Error Resume Next
            '     Set fso = CreateObject("Scripting.FileSystemObject" )
            '         fso.DeleteFile(Server.MapPath(str_file_path& DosyaAc("FOTOGRAF") &"")) 
            '         fso.DeleteFile(Server.MapPath(str_file_path&"M_"& DosyaAc("FOTOGRAF") &""))
            '         fso.DeleteFile(Server.MapPath(str_file_path&"T_"& DosyaAc("FOTOGRAF") &"")) 
            '         fso.DeleteFile(Server.MapPath(str_file_path&"cms_"& DosyaAc("FOTOGRAF") &""))
            '         fso.DeleteFile(Server.MapPath(str_file_path&"orj_"& DosyaAc("FOTOGRAF") &"")) 
            '     Set Fso = Nothing       
            '     On Error Goto 0
            '     Conn.Execute("DELETE FROM tbl_urun_fotograflar WHERE ID = " & DosyaAc("ID") & "")
            ' DosyaAc.MoveNext : Loop
            ' DosyaAc.Close : Set DosyaAc = Nothing   
            Verbose "DOWNLOAD", "Dosya İndirilmiş, Pas geçiliyor..."
            DownloadFile = True 
            Exit Property 
        End If

        Set objHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
            objHTTP.Open "GET", FromURL
            objHTTP.Send

            If objHTTP.Status = 200 Then
                Select Case DOWNLOAD_WITH
                    Case "ASPJPEG"
                        Verbose "DOWNLOAD", ""& FromURL &" İndiriliyor..."
                        
                        Set Jpeg = Server.CreateObject("Persits.Jpeg")
                            Jpeg.OpenBinary( objHTTP.responseBody )
                            Jpeg.Save DOWNLOAD_FILE_PATH & FileName
                        Set Jpeg = Nothing
                        
                        Verbose "DOWNLOAD", ""& FromURL &" İndirildi"

                        CopyOriginal DOWNLOAD_FILE_PATH, FileName
                        
                        DownloadFile = True
                    Case "ADODB"
                        Verbose "DOWNLOAD", ""& FromURL &" İndiriliyor..."
                        
                        Set objADOStream = CreateObject("ADODB.Stream")
                            objADOStream.Open
                            objADOStream.Type = 1
                            objADOStream.Write objHTTP.responseBody
                            objADOStream.Position = 0
                            objADOStream.SaveToFile DOWNLOAD_FILE_PATH & FileName, YesOverWrite
                            objADOStream.Close
                        Set objADOStream = Nothing
                        
                        Verbose "DOWNLOAD", ""& FromURL &" İndirildi"
                        
                        CopyOriginal DOWNLOAD_FILE_PATH, FileName
                        
                        DownloadFile = True
                    Case Else 
                        Verbose "DOWNLOAD", "Kayıt İstemcisi Hatası ("& DOWNLOAD_WITH &")"
                        DownloadFile = False
                End Select
                ' If DownloadFile = True Then 
                '     SetPhotos DOWNLOAD_FILE_PATH, FileName
                ' End If
            Else
                Verbose "DOWNLOAD", ""& objHTTP.responseText &""
            End If
        Set objHTTP = nothing
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get CopyOriginal(FilePath, FileName)
        If Not PROTECT_ORIGINAL_FILE = True Then 
            Exit Property
        End If

        Set Fs = Server.CreateObject("Scripting.FileSystemObject")
        ' Set F = Fs.GetFile(FilePath)
        '     FileSize = F.Size
        ' Set F = Nothing
            Fs.CopyFile FilePath, (FilePath & "NONTINIFIED_"&FileName), True
        Set Fs = Nothing
        
        Verbose "IMAGE PROTECT", ""& FileName &" Orjinali Korunuyor"
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Function BuildSQL(dicObj, dbId)
        Dim keyz, itemz
        keyz = dicObj.keys
        itemz= dicObj.items

        If dicObj.Count= 0 Then 
            BuildSQL = False
            Exit Function
        End If

        Dim sqlString
            sqlString = "UPDATE tbl_urun SET " 

        For i=0 To dicObj.Count -1
            sqlString = sqlString & keyz(i) &"='"& itemz(i) &"', "
        Next

        BuildSQL = Left(sqlString, Len(sqlString)-2) & " WHERE ID="& dbId &""
    End Function
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get Verbose(Title, Text)
        If DEBUG = False Then Exit Property
        
        Response.Write "....["& Title &"] : "& Text &"<br />"
        Response.Flush
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Property Get SetPhotos(FilePath, FileName)
        Call ImageUploadSettings()

        kodd = kod(15, "" )
        Dosya = Kodd & FileName         

        ' If IsFileExist(DOWNLOAD_FILE_PATH & FileName) = True Then 
        '     Verbose "SETPHOTOS", "Dosya İndirilmiş, Pas geçiliyor..."
        '     SetPhotos = "" 
        '     Exit Property 
        ' End If

        ' Crop CMS Boyutu
        '------------------------------------------------------------

        Set Jpeg1 = Server.CreateObject("Persits.Jpeg")
            Jpeg1.Open DOWNLOAD_FILE_PATH & FileName
            Jpeg1.Quality       = 60
            SourceAspectRatio   = Jpeg1.Width / Jpeg1.Height
            DesiredAspectRatio  = cms_image_square_width / cms_image_square_height
            if SourceAspectRatio > DesiredAspectRatio then
                Jpeg1.Height    = cms_image_square_height
                Jpeg1.Width     = cms_image_square_height * SourceAspectRatio
            Else 
                Jpeg1.Width     = cms_image_square_width
                Jpeg1.Height    = cms_image_square_width / SourceAspectRatio
            End If
            X0 = (Jpeg1.Width - cms_image_square_width) / 2
            Y0 = (Jpeg1.Height - cms_image_square_height) / 2
            X1 = X0 + cms_image_square_width
            Y1 = Y0 + cms_image_square_height
            Jpeg1.Crop X0, Y0, X1, Y1
            Jpeg1.Save DOWNLOAD_FILE_PATH & "cms_"& Dosya
        Set Jpeg1 = Nothing
        ' Verbose "IMAGE", "Crop Boyutu Tamam"
        ' Sleep 1

        ' Thumbnail CMS Boyutu
        '------------------------------------------------------------
        OpenFile            = DOWNLOAD_FILE_PATH & FileName
        SaveFile            = DOWNLOAD_FILE_PATH & "T_" & Dosya
        CompressQuality     = THUMBNAIL_QUALITY
        MaxWidth            = THUMBNAIL_WIDTH
        SetFit              = THUMBNAIL_FIT
        ASPJpeg_ReSize OpenFile, SaveFile, CompressQuality, MaxWidth, SetFit
        ' Verbose "IMAGE", "Thumbnail Boyutu Tamam"
        ' Sleep 1

        ' Medium Boyut
        '------------------------------------------------------------
        OpenFile            = DOWNLOAD_FILE_PATH & FileName
        SaveFile            = DOWNLOAD_FILE_PATH & "M_" & Dosya
        CompressQuality     = MEDIUM_QUALITY
        MaxWidth            = MEDIUM_WIDTH
        SetFit              = MEDIUM_FIT
        ASPJpeg_ReSize OpenFile, SaveFile, CompressQuality, MaxWidth, SetFit
        ' Verbose "IMAGE", "Medium Boyutu Tamam"
        ' Sleep 1
        
        ' Large Boyut
        '------------------------------------------------------------
        OpenFile            = DOWNLOAD_FILE_PATH & FileName
        SaveFile            = DOWNLOAD_FILE_PATH & Dosya
        CompressQuality     = FULL_QUALITY
        MaxWidth            = FULL_WIDTH
        SetFit              = FULL_FIT
        ASPJpeg_ReSize OpenFile, SaveFile, CompressQuality, MaxWidth, SetFit
        ' Verbose "IMAGE", "Large Boyutu Tamam"
        ' Sleep 1

        Verbose "IMAGE", "İmajlar Tamamlandı"
        SetPhotos = Dosya
    End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
    Public Function GetFileName(FileURL)
        If Len(FileURL) < 5 Then 
            GetFileName = False
            Exit Function
        End If

        Dim FileName
            FileName = Split(FileURL, "/")
        GetFileName = FileName( UBound(FileName) )
    End Function
End Class 
%>

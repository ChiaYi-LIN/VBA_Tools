Attribute VB_Name = "Dictionary"
Public PUBCompanyName, PUBLocationName, PUBGIDRef, PUBInformation, PUBOIDRef, PUBLevelHeader, PUBParentHeader As Range
Public PUBLocationLevel As String
Public theElectronic, theTechnology, theSystem, theScience, theEngineer, theAutomation, theEnterprise As Variant
Public PROGRESS_BAR_MODE As Integer

Sub DictionaryInitialization()

theElectronic = Array(" ELECTRONIK ", " ELECTRONIS ", " ELECTRONISCHE ", " ELECTRONICA ", " ELECTRONICAL ", _
    " ELECTRONICAS ", " ELECTRIQ ", " ELECTRIQUE ", " ELECTRONIQ ", " ELECTRONIQU ", " ELECTRONIQUE ", " ELECTRONIQUEJM ", _
    " ELECTRONIQUES ", " ELEKTRONIK ", " ELECTRONICS ", " ELECTRONIC ", " ELETTRONICA ", " ELECRTONICS ", " ELECTROC ", _
    " ELECTRIC ", " ELEKTRIK ", " ELECTRICITY ", " ELECTRON ", " ELEKTRONISK ", " ELEKTRONISKI ", " ELEKTRICITET ", _
    " ELEKTRONIKO ", " ELEKTRIZITATEA ", " LISTRIK ", " ELEKTRONIKA ", " ELEKTRO ", " ELECTRICIDADE ", " ELECTRICITAT ", _
    " ELECTRICITATE ", " ELEKTRONIESE ", " ELEKTRISITEIT ", "ELEKTROANYSK ", " ELEKTRONINIS ", " ELEKTRA ", " ELETRIK ", _
    " ELEKTRONIKUS ", " ELEKTROMOSSAG ", " ELECTRONICO ", " ELECTRICIDAD ", " ELEKTRON ", " ZAMAGETSI ", " MAGETSI ", _
    " ELECTRICAE ", " ELEKTRIBA ", " ELEKTRONSKI ", " STRUJA ", " ELEKTRONICZNY ", " ELEKTRYCZNOSC ", " ELEKTRONINEN ", _
    " NGEKHOMPYUTHA ", " UMBANE ", " UGESI ", " ELECTRONIG ", " TRYDAN ", " ELETTRONICU ", " HLUAV TAWS XOB ", " ELATRIK ", _
    " ELEKTRISITET ", " ELEKTWONIK ", " ELEKTRISITE ", " ELEKTR ", " ENERGIYASI ", " ELEKTAROONIK ", " KORONTO ", " ELETTRONIKU ", _
    " ELETTRIKU ", " HERINARATRA ", " KURYENTE ", " ELEKTRONICKA ", " ELEKTRONICKY ", " ELEKTRONISCH ", " ELEKTRICITEIT ", _
    " UMEME ", " ELEKTRINA ", " ELEKTRONSKO ", " ELEKTRIKA ", " ELEKTROONILINE ", " ELEKTRIT ", " LEICTREONACH ", _
    " LEICTREACHAS ", " ELEKTRONIKE ", " MOTLAKASE ", " ELETTRONICO ", " ELETTRICITA ", " ELETRONICO ", " ELETRICIDADE ", _
    " LANTARKI ", " WUTAR ", " ELEKTRONESCH ", " STROUM ", " ELETORONI ", " ELETISE ", " DEALANACH ", " DEALAN ")
    
theTechnology = Array(" TECHNOLOGY ", " TECHNICAL ", " TECHNOLOGIES ", " TECHNIQUE ", " TEKNOLOJI ", " TEKNOLOJILER ", _
    " TEKNIK ", " TEKNOLOGI ", " TEKNOLOGIER ", " TEKNISK ", " TEKNOLOGIA ", " TEKNIKOAK ", " TEKNOLOGIO ", " TEKNOLOGIOJ ", _
    " TEKNIKA ", " TECNOLOXIA ", " TECNOLOXIAS ", " TECNOLOGIA ", " TECNIC ", " TEGNOLOGIE ", " TEGNIESE ", " TECHNOLOGYEN ", _
    " TECHNYSK ", " TECHNOLOGIJA ", " TECHNOLOGIJAS ", " TECHNINIS ", " NKA NA UZU ", " TEKNUZU ", " TECHNOLOGIAK ", " MUSZAKI ", _
    " TEKNIS ", " TECNOLOGIAS ", " TECNICO ", " TEHNOLOGIJA ", " TEHNOLOGIJE ", " TEHNICKA ", " TEXNOLOGIYA ", " TEXNOLOGIYALARI ", _
    " TEXNIKI ", " MAKANEMA ", " ZAMAKONO ", " TECHNOLOGIAE ", " TEHNOLOGIJAS ", " TEHNISKS ", " LA TECHNOLOGIE ", _
    " LES TECHNOLOGIES ", " TEHNICKI ", " TECHNOLOGIA ", " TECHNOLOGIE ", " TECHNICZNY ", " TEKNIIKKA ", " TEKNOLOGIOIDEN ", _
    " TEKNINEN ", " TEKNOLOGJI ", " TEKNOLOGJITE ", " UBUCHWEPHESHA ", " ZOBUGCISA ", " UBUCHWEPHESHE ", " TECHNOLEG ", _
    " KEV PAUB ", " TEKNOLOCI ", " TEKNOLOJIYEN ", " TEKNIKI ", " TEXNIK ", " TIKNOOLAJI ", " TEKNOOLOOJIYADA ", " FARSAMO ", _
    " TEKNOLOGIJA ", " TEKNOLOGIJI ", " TEKNIKU ", " TEKNIKAL ", " TEKNOLOJIA ", " ARA-TEKNIKA ", " TEKNOLOHIYA ", " TECHNIKA ", _
    " TECHNOLOGII ", " TECHNICKY ", " TECHNOLOGIEEN ", " TECHNISCH ", " KIUFUNDI ", " TEHNOLOGIJO ", " TEHNICNO ", " TEHNOLOOGIA ", _
    " TEHNOLOOGIAD ", " TEHNILINE ", " TECHNOLEGAU ", " TECHNEGOL ", " TECNULUGIA ", " TECNULUGII ", " TECNICU ", _
    " TEICNEOLAIOCHT ", " TEICNEOLAIOCHTAI ", " TEICNIULA ", " THEKNOLOJI ", " TSA SETSEBI ", " TECNOLOGIE ", " FASAHA ", _
    " TECHNESCH ", " TEKONOLOSI ", " TEKINOLOSI ", " TEHNOLOGIE ", " TEHNOLOGII ", " TEHNIC ", " TEICNEOLAS ", " TEICNEOLASAN ", _
    " TEICNIGEACH ", " TECHNOS ")
    
theSystem = Array(" SYSTEM ", " SYSTEMS ", " SISTEMA ", " SYSTEEM ", " USORO ", " SISTEMI ", " RAFITRA ", " TSARIN ", _
    " FAIGA ", " SISTEMAS ", " SYSTEMEN ", " SYSTEMER ", " MGA SISTEMA ", " SISTEMETAN ", " SISTEMOJ ", " STELSELS ", _
    " RENDSZEREK ", " MACHITIDWE ", " JARJESTELMAT ", " IINKQUBO ", " IZINHLELO ", " PERGALEN ", " SISTEM YO ", " TIZIMLARI ", _
    " NIDAAMYADA ", " MIFUMO ", " SUSTEEMID ", " CORAIS ", " SIOSTAMAN ", " SISTEMLER ", " SISTEMES ", " SISTEMOS ", _
    " SUSTAVI ", " SISTEMLERI ", " SYSTEMY ", " SISTEMET ", " SYSTEMAU ", " ZVIRONGWA ", " SISTEMOV ", " LITSAMAISO ", _
    " SISTEME ", " SISTEMAN ", " SISTEMO ", " STELSEL ", " RENDSZER ", " SISTIM ", " DONGOSOLO ", " RATIO ", " SYSTEME ", _
    " JARJESTELMA ", " INKQUBO ", " UHLELO ", " KAW LUS ", " TIZIMI ", " NIDAAMKA ", " MFUMO ", " SUSTEEM ", " CORAS ", _
    " SYSTEMET ", " SIOSTAM ")
    
theScience = Array(" SCIENCE ", " SCIENCES ", " BILIM ", " VIDENSKAB ", " ZIENTZIA ", " ILMU ", " SCIENCO ", " WETENSKAP ", _
    " WITTENSKIP ", " MOKSLAS ", " TUDOMANY ", " ELMU PANGAWERUH ", " ELM ", " SCIENTIA ", " ZINATNE ", " TIEDE ", " SHKENCE ", _
    " YESAYENSI ", " GWYDDONIAETH ", " VITENSKAP ", " ILM-FAN ", " IX-XJENZA ", " SIYENSIYA ", " SAYENZI ", " WETENSCHAP ", _
    " TEADUS ", " EOLAIOCHT ", " VETENSKAP ", " KIMIYYA ", " WESSENSCHAFT ", " STIINTA ", " SAIDHEANS ", " NAUKA ", " ISAYENSI ", _
    " ZANIST ", " SYANS ", " SAYNISKA ", " SAINS ", " SIANSA ", " VEDA ", " SAYANSI ", " ZNANOST ", " AGHAM ", " SAENSE ", _
    " SCIENZA ", " CIENCIA ", " SAIENISI ", " BILIMLER ", " VIDENSKABER ", " ZIENTZIAK ", " ELMU ", " SCIENCOJ ", " CIENCIES ", _
    " WETENSKAPPE ", " WITTENSKIPPEN ", " MOKSLAI ", " TUDOMANYOK ", " ELML?R ", " ARTIUM ", " ZINATNES ", " LES SCIENCES ", _
    " NAUKE ", " NAUKI ", " SHKENCAT ", " ZE NZU LULWAZI ", " GWYDDORAU ", " VITENSKAPER ", " ILMLAR ", " XJENZI ", _
    " MGA SIYENSIYA ", " WETENSCHAPPEN ", " TEADUSED ", " NA HEOLAIOCHTAI ", " VETENSKAPER ", " KIMIYYAR ", " WESSENSCHAFTEN ", _
    " STIINTE ", " SAIDHEANSAN ", " CIENCIAS ", " SAYENS? ", " ZNANOSTI ", " SCIENZE ", " VEDY ")
    
theEngineer = Array(" ENGINEER ", " ENGINEERING ", " MUHENDIS ", " INGENIOR ", " ENXENEIRO ", " INGENIEUR ", _
    " YNGENIEUR ", " INSINYUR ", " INGENIERO ", " INZENJER ", " INJENIERA ", " INSENER ", " INJINIYA ", " MUHENDISLIK ", _
    " INZENJERING ", " INGENIARIA ", " ENGINYER ", " INZINIERIUS ", " MERNOK ", " FECTUM ", " INZENIERIS ", " INZYNIER ", _
    " INSINOORI ", " INXHINIER ", " NJINELI ", " UNJINIYELA ", " PEIRIANNYDD ", " HENDESE ", " ENJENYE ", " MUHANDIS ", _
    " INJINEER ", " INGINIER ", " JURUTERA ", " INZENYR ", " MHANDISI ", " INZINIER ", " INZENIR ", " INNEALTOIR ", _
    " INGENJOR ", " MOENJINIERE ", " INGEGNERE ", " ENGENHEIRO ", " INISINIA ", " INGINER ", " INNLEADAIR ", " INGENIARITZA ", _
    " ENGINYERIA ", " INZINERIJA ", " MERNOKI ", " IPSUM ", " INZENIERIJA ", " INZYNIERIA ", " TEKNIIKKA ", " INXHINIERI ", _
    " ZOBUNJINELI ", " UBUNJINIYELA ", " PEIRIANNEG ", " ENDAZYARIYE ", " JENI ", " MUHANDISLIK ", " INJINEERNIMADA ", _
    " INGINERIJA ", " KEJURUTERAAN ", " INZENYRSTVI ", " UHANDISI ", " STROJARSTVO ", " INZENIRING ", " INNEALTOIREACHT ", _
    " BOENJINIERE ", " INGEGNERIA ", " ENGENHARIA ", " ENISINIA ", " INGINERIE ", " INNLEADAIREACHD ", " INGENIERISTIKO ", _
    " INGENIEURSWESE ", " INJINIA ", " REKAYASA ", " INGENIERIA ", " INGENIERIE ", " INGIGNIRIA ", " PROSJEKTERING ", _
    " BOUWKUNDE ", " AIKIN INJINIYA ", " DEIFBAU ")

theAutomation = Array(" AUTOMATION ", " AUTOMATISERING ", " OTOMATISASI ", " AUTOMATIZACION ", " AUTOMATIZACIJA ", _
    " AUTOMATISATION ", " OTOMASYON ", " AUTOMATIZAZIOA ", " AUTOMATIGO ", " AUTOMATITZACIO ", " OUTOMATISERING ", _
    " AUTOMATISEARRING ", " AUTOMATIKA ", " AKPAAKA ", " AUTOMATIZALAS ", " HAL NU NGAJADIKEUN OTOMATIS ", " AVTOMATLASDIRMA ", _
    " ZOKHA ", " AUTOMATYZACJA ", " AUTOMAATIO ", " AUTOMATIZIM ", " NGOKUZENZEKELAYO ", " UKUZENZEKELAYO ", " AWTOMEIDDIO ", _
    " AUTOMATIZAZIONE ", " OTOMATIKI ", " AUTOMASJON ", " AVTOMATLASHTIRISH ", " AWTOMATIZZAZZJONI ", " AUTOMASI ", " AUTOMATIQUE ", _
    " AUTOMATIZACE ", " AUTOMATIZACIA ", " AVTOMATIZACIJA ", " AUTOMATISEERIMINE ", " UATHOIBRIU ", " HO IKETSETSA LETHO ", _
    " AUTOMAZIONE ", " AUTOMACAO ", " AIKI DA KAI ", " AUTOMATISATIOUN ", " AUTOMATIZARE ", " FEIN-GHLUASAD ", _
    " AUTOMOTIVE ", " OTOMOTIV ", " AUTOMOBILGINTZA ", " AUTOMOBIL ", " AUTOMOCIO ", " MOTOR ", " AUTOMOBILIAI ", _
    " AKURUNGWA ", " AUTOIPARI ", " OTOMOTIF ", " AUTOMOTOR ", " AUTOMOBILSKI ", " AVTOMOBIL ", " MAGALIMOTO ", " EGET ", _
    " AUTOMOBILU RUPNIECIBA ", " AUTOMOBILE ", " AUTOMOBILSKA ", " AUTOMOBILOWY ", " AUTOJEN ", " AUTOMOBILISTIK ", _
    " IMOTO ", " IZIMOTO ", " MODUROL ", " AUTOMOBILISTICA ", " TSHEB ", " OTOTOTIK ", " OTOMOBIL ", " BAABUURTA ", _
    " TAL-KAROZZI ", " AUTOMOTIF ", " FIARA ", " AUTOMOBILOVY PRUMYSL ", " MOTOKARI ", " MAGARI ", " AVTOMOBILSKA INDUSTRIJA ", _
    " AUTOTOOSTUS ", " FEITHICLEACH ", " BIL ", " LIKOLOI ", " SETTORE AUTOMOBILISTICO ", " AUTOMOTIVO ", " MOTA ", " TAAVALE AFI ", _
    " UIDHEAMACHD ")
    
theEnterprise = Array(" ENTERPRISE ", " KURULUS ", " FORETAGENDE ", " ENPRESA ", " PERUSAHAAN ", " ENTREPRENO ", " EMPRESA ", _
    " ONDERNEMING ", " BEDRIUW ", " ?MONE ", " ULO ORU ", " VALLALKOZAS ", " PAUSAHAAN ", " PODUZECE ", " MUASSISA ", " MALONDA ", _
    " COEPTIS ", " UZNEMUMS ", " ENTREPRISE ", " PREDUZECE ", " PRZEDSIEBIORSTWO ", " YRITYS ", " NDERMARRJE ", " SHISHINI ", _
    " IBHIZINISI ", " MENTER ", " IMPRESE ", " SIRKET ", " BEDRIFTEN ", " ANTREPRIZ ", " KORXONA ", " SHIRKAD ", " INTRAPRIZA ", _
    " ORINASA ", " NEGOSYO ", " PODNIK ", " BUSINESS ", " BIASHARA ", " PODJETJE ", " ETTEVOTE ", " FIONTAR ", " FORETAG ", " KHOEBO ", _
    " IMPRESA ", " EMPREENDIMENTO ", " CINIKI ", " ATINAE ", " AFACERE ", " IOMAIRT ")

''''''' " AUTO " should not be in the array
'theAutomation = Array(" AUTOMATION ", " AUTOMATISERING ", " OTOMATISASI ", " AUTOMATIZACION ", " AUTOMATIZACIJA ", _
    " AUTOMATISATION ", " OTOMASYON ", " AUTOMATIZAZIOA ", " AUTOMATIGO ", " AUTOMATITZACIO ", " OUTOMATISERING ", _
    " AUTOMATISEARRING ", " AUTOMATIKA ", " AKPAAKA ", " AUTOMATIZALAS ", " HAL NU NGAJADIKEUN OTOMATIS ", " AVTOMATLASDIRMA ", _
    " ZOKHA ", " AUTOMATYZACJA ", " AUTOMAATIO ", " AUTOMATIZIM ", " NGOKUZENZEKELAYO ", " UKUZENZEKELAYO ", " AWTOMEIDDIO ", _
    " AUTOMATIZAZIONE ", " OTOMATIKI ", " AUTOMASJON ", " AVTOMATLASHTIRISH ", " AWTOMATIZZAZZJONI ", " AUTOMASI ", " AUTOMATIQUE ", _
    " AUTOMATIZACE ", " AUTOMATIZACIA ", " AVTOMATIZACIJA ", " AUTOMATISEERIMINE ", " UATHOIBRIU ", " HO IKETSETSA LETHO ", _
    " AUTOMAZIONE ", " AUTOMACAO ", " AIKI DA KAI ", " AUTOMATISATIOUN ", " AUTOMATIZARE ", " FEIN-GHLUASAD ", _
    " AUTOMOTIVE ", " OTOMOTIV ", " AUTOMOBILGINTZA ", " AUTO ", " AUTOMOBIL ", " AUTOMOCIO ", " MOTOR ", " AUTOMOBILIAI ", _
    " AKURUNGWA ", " AUTOIPARI ", " OTOMOTIF ", " AUTOMOTOR ", " AUTOMOBILSKI ", " AVTOMOBIL ", " MAGALIMOTO ", " EGET ", _
    " AUTOMOBILU RUPNIECIBA ", " AUTOMOBILE ", " AUTOMOBILSKA ", " AUTOMOBILOWY ", " AUTOJEN ", " AUTOMOBILISTIK ", _
    " IMOTO ", " IZIMOTO ", " MODUROL ", " AUTOMOBILISTICA ", " TSHEB ", " OTOTOTIK ", " OTOMOBIL ", " BAABUURTA ", _
    " TAL-KAROZZI ", " AUTOMOTIF ", " FIARA ", " AUTOMOBILOVY PRUMYSL ", " MOTOKARI ", " MAGARI ", " AVTOMOBILSKA INDUSTRIJA ", _
    " AUTOTOOSTUS ", " FEITHICLEACH ", " BIL ", " LIKOLOI ", " SETTORE AUTOMOBILISTICO ", " AUTOMOTIVO ", " MOTA ", " TAAVALE AFI ", _
    " UIDHEAMACHD ")
        
End Sub

Function backup()
    Dim theElectronic As Variant
    Dim theTechnology As Variant
    Dim theSystem As Variant
    Dim theScience As Variant
    Dim theEngineer As Variant
    Dim theAutomation As Variant
    Dim theEnterprise As Variant

    'Standardize certain common words - ELECTR
    theElectronic = Array(" ELECTRONIK ", " ELECTRONIS ", " ELECTRONISCHE ", " ELECTRONICA ", " ELECTRONICAL ", _
    " ELECTRONICAS ", " ELECTRIQ ", " ELECTRIQUE ", " ELECTRONIQ ", " ELECTRONIQU ", " ELECTRONIQUE ", " ELECTRONIQUEJM ", _
    " ELECTRONIQUES ", " ELEKTRONIK ", " ELECTRONICS ", " ELECTRONIC ", " ELETTRONICA ", " ELECRTONICS ", " ELECTROC ", _
    " ELECTRIC ", " ELEKTRIK ", " ELECTRICITY ", " ELECTRON ", " ELEKTRONISK ", " ELEKTRONISKI ", " ELEKTRICITET ", _
    " ELEKTRONIKO ", " ELEKTRIZITATEA ", " LISTRIK ", " ELEKTRONIKA ", " ELEKTRO ", " ELECTRICIDADE ", " ELECTRICITAT ", _
    " ELECTRICITATE ", " ELEKTRONIESE ", " ELEKTRISITEIT ", "ELEKTROANYSK ", " ELEKTRONINIS ", " ELEKTRA ", " ELETRIK ", _
    " ELEKTRONIKUS ", " ELEKTROMOSSAG ", " ELECTRONICO ", " ELECTRICIDAD ", " ELEKTRON ", " ZAMAGETSI ", " MAGETSI ", _
    " ELECTRICAE ", " ELEKTRIBA ", " ELEKTRONSKI ", " STRUJA ", " ELEKTRONICZNY ", " ELEKTRYCZNOSC ", " ELEKTRONINEN ", _
    " NGEKHOMPYUTHA ", " UMBANE ", " UGESI ", " ELECTRONIG ", " TRYDAN ", " ELETTRONICU ", " HLUAV TAWS XOB ", " ELATRIK ", _
    " ELEKTRISITET ", " ELEKTWONIK ", " ELEKTRISITE ", " ELEKTR ", " ENERGIYASI ", " ELEKTAROONIK ", " KORONTO ", " ELETTRONIKU ", _
    " ELETTRIKU ", " HERINARATRA ", " KURYENTE ", " ELEKTRONICKA ", " ELEKTRONICKY ", " ELEKTRONISCH ", " ELEKTRICITEIT ", _
    " UMEME ", " ELEKTRINA ", " ELEKTRONSKO ", " ELEKTRIKA ", " ELEKTROONILINE ", " ELEKTRIT ", " LEICTREONACH ", _
    " LEICTREACHAS ", " ELEKTRONIKE ", " MOTLAKASE ", " ELETTRONICO ", " ELETTRICITA ", " ELETRONICO ", " ELETRICIDADE ", _
    " LANTARKI ", " WUTAR ", " ELEKTRONESCH ", " STROUM ", " ELETORONI ", " ELETISE ", " DEALANACH ", " DEALAN ")
    
    For i = 0 To UBound(theElectronic)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theElectronic(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theElectronic(i), " ELECTR ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theElectronic(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - Tech
    theTechnology = Array(" TECHNOLOGY ", " TECHNICAL ", " TECHNOLOGIES ", " TECHNIQUE ", " TEKNOLOJI ", " TEKNOLOJILER ", _
    " TEKNIK ", " TEKNOLOGI ", " TEKNOLOGIER ", " TEKNISK ", " TEKNOLOGIA ", " TEKNIKOAK ", " TEKNOLOGIO ", " TEKNOLOGIOJ ", _
    " TEKNIKA ", " TECNOLOXIA ", " TECNOLOXIAS ", " TECNOLOGIA ", " TECNIC ", " TEGNOLOGIE ", " TEGNIESE ", " TECHNOLOGYEN ", _
    " TECHNYSK ", " TECHNOLOGIJA ", " TECHNOLOGIJAS ", " TECHNINIS ", " NKA NA UZU ", " TEKNUZU ", " TECHNOLOGIAK ", " MUSZAKI ", _
    " TEKNIS ", " TECNOLOGIAS ", " TECNICO ", " TEHNOLOGIJA ", " TEHNOLOGIJE ", " TEHNICKA ", " TEXNOLOGIYA ", " TEXNOLOGIYALARI ", _
    " TEXNIKI ", " MAKANEMA ", " ZAMAKONO ", " TECHNOLOGIAE ", " TEHNOLOGIJAS ", " TEHNISKS ", " LA TECHNOLOGIE ", _
    " LES TECHNOLOGIES ", " TEHNICKI ", " TECHNOLOGIA ", " TECHNOLOGIE ", " TECHNICZNY ", " TEKNIIKKA ", " TEKNOLOGIOIDEN ", _
    " TEKNINEN ", " TEKNOLOGJI ", " TEKNOLOGJITE ", " UBUCHWEPHESHA ", " ZOBUGCISA ", " UBUCHWEPHESHE ", " TECHNOLEG ", _
    " KEV PAUB ", " TEKNOLOCI ", " TEKNOLOJIYEN ", " TEKNIKI ", " TEXNIK ", " TIKNOOLAJI ", " TEKNOOLOOJIYADA ", " FARSAMO ", _
    " TEKNOLOGIJA ", " TEKNOLOGIJI ", " TEKNIKU ", " TEKNIKAL ", " TEKNOLOJIA ", " ARA-TEKNIKA ", " TEKNOLOHIYA ", " TECHNIKA ", _
    " TECHNOLOGII ", " TECHNICKY ", " TECHNOLOGIEEN ", " TECHNISCH ", " KIUFUNDI ", " TEHNOLOGIJO ", " TEHNICNO ", " TEHNOLOOGIA ", _
    " TEHNOLOOGIAD ", " TEHNILINE ", " TECHNOLEGAU ", " TECHNEGOL ", " TECNULUGIA ", " TECNULUGII ", " TECNICU ", _
    " TEICNEOLAIOCHT ", " TEICNEOLAIOCHTAI ", " TEICNIULA ", " THEKNOLOJI ", " TSA SETSEBI ", " TECNOLOGIE ", " FASAHA ", _
    " TECHNESCH ", " TEKONOLOSI ", " TEKINOLOSI ", " TEHNOLOGIE ", " TEHNOLOGII ", " TEHNIC ", " TEICNEOLAS ", " TEICNEOLASAN ", _
    " TEICNIGEACH ", " TECHNOS ")
    For i = 0 To UBound(theTechnology)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theTechnology(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theTechnology(i), " TECH ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theTechnology(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - SYS
    theSystem = Array(" SYSTEM ", " SYSTEMS ", " SISTEMA ", " SYSTEEM ", " USORO ", " SISTEMI ", " RAFITRA ", " TSARIN ", _
    " FAIGA ", " SISTEMAS ", " SYSTEMEN ", " SYSTEMER ", " MGA SISTEMA ", " SISTEMETAN ", " SISTEMOJ ", " STELSELS ", _
    " RENDSZEREK ", " MACHITIDWE ", " JARJESTELMAT ", " IINKQUBO ", " IZINHLELO ", " PERGALEN ", " SISTEM YO ", " TIZIMLARI ", _
    " NIDAAMYADA ", " MIFUMO ", " SUSTEEMID ", " CORAIS ", " SIOSTAMAN ", " SISTEMLER ", " SISTEMES ", " SISTEMOS ", _
    " SUSTAVI ", " SISTEMLERI ", " SYSTEMY ", " SISTEMET ", " SYSTEMAU ", " ZVIRONGWA ", " SISTEMOV ", " LITSAMAISO ", _
    " SISTEME ", " SISTEMAN ", " SISTEMO ", " STELSEL ", " RENDSZER ", " SISTIM ", " DONGOSOLO ", " RATIO ", " SYSTEME ", _
    " JARJESTELMA ", " INKQUBO ", " UHLELO ", " KAW LUS ", " TIZIMI ", " NIDAAMKA ", " MFUMO ", " SUSTEEM ", " CORAS ", _
    " SYSTEMET ", " SIOSTAM ")
    For i = 0 To UBound(theSystem)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theSystem(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theSystem(i), " SYS ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theSystem(i), 1)
    Wend
    Next i
        
    'Standardize certain common words - SCI
    theScience = Array(" SCIENCE ", " SCIENCES ", " BILIM ", " VIDENSKAB ", " ZIENTZIA ", " ILMU ", " SCIENCO ", " WETENSKAP ", _
    " WITTENSKIP ", " MOKSLAS ", " TUDOMANY ", " ELMU PANGAWERUH ", " ELM ", " SCIENTIA ", " ZINATNE ", " TIEDE ", " SHKENCE ", _
    " YESAYENSI ", " GWYDDONIAETH ", " VITENSKAP ", " ILM-FAN ", " IX-XJENZA ", " SIYENSIYA ", " SAYENZI ", " WETENSCHAP ", _
    " TEADUS ", " EOLAIOCHT ", " VETENSKAP ", " KIMIYYA ", " WESSENSCHAFT ", " STIINTA ", " SAIDHEANS ", " NAUKA ", " ISAYENSI ", _
    " ZANIST ", " SYANS ", " SAYNISKA ", " SAINS ", " SIANSA ", " VEDA ", " SAYANSI ", " ZNANOST ", " AGHAM ", " SAENSE ", _
    " SCIENZA ", " CIENCIA ", " SAIENISI ", " BILIMLER ", " VIDENSKABER ", " ZIENTZIAK ", " ELMU ", " SCIENCOJ ", " CIENCIES ", _
    " WETENSKAPPE ", " WITTENSKIPPEN ", " MOKSLAI ", " TUDOMANYOK ", " ELML?R ", " ARTIUM ", " ZINATNES ", " LES SCIENCES ", _
    " NAUKE ", " NAUKI ", " SHKENCAT ", " ZE NZU LULWAZI ", " GWYDDORAU ", " VITENSKAPER ", " ILMLAR ", " XJENZI ", _
    " MGA SIYENSIYA ", " WETENSCHAPPEN ", " TEADUSED ", " NA HEOLAIOCHTAI ", " VETENSKAPER ", " KIMIYYAR ", " WESSENSCHAFTEN ", _
    " STIINTE ", " SAIDHEANSAN ", " CIENCIAS ", " SAYENS? ", " ZNANOSTI ", " SCIENZE ", " VEDY ")
    For i = 0 To UBound(theScience)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theScience(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theScience(i), " SCI ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theScience(i), 1)
    Wend
    Next i
     
    'Standardize certain common words - ENG
    theEngineer = Array(" ENGINEER ", " ENGINEERING ", " MUHENDIS ", " INGENIOR ", " ENXENEIRO ", " INGENIEUR ", _
    " YNGENIEUR ", " INSINYUR ", " INGENIERO ", " INZENJER ", " INJENIERA ", " INSENER ", " INJINIYA ", " MUHENDISLIK ", _
    " INZENJERING ", " INGENIARIA ", " ENGINYER ", " INZINIERIUS ", " MERNOK ", " FECTUM ", " INZENIERIS ", " INZYNIER ", _
    " INSINOORI ", " INXHINIER ", " NJINELI ", " UNJINIYELA ", " PEIRIANNYDD ", " HENDESE ", " ENJENYE ", " MUHANDIS ", _
    " INJINEER ", " INGINIER ", " JURUTERA ", " INZENYR ", " MHANDISI ", " INZINIER ", " INZENIR ", " INNEALTOIR ", _
    " INGENJOR ", " MOENJINIERE ", " INGEGNERE ", " ENGENHEIRO ", " INISINIA ", " INGINER ", " INNLEADAIR ", " INGENIARITZA ", _
    " ENGINYERIA ", " INZINERIJA ", " MERNOKI ", " IPSUM ", " INZENIERIJA ", " INZYNIERIA ", " TEKNIIKKA ", " INXHINIERI ", _
    " ZOBUNJINELI ", " UBUNJINIYELA ", " PEIRIANNEG ", " ENDAZYARIYE ", " JENI ", " MUHANDISLIK ", " INJINEERNIMADA ", _
    " INGINERIJA ", " KEJURUTERAAN ", " INZENYRSTVI ", " UHANDISI ", " STROJARSTVO ", " INZENIRING ", " INNEALTOIREACHT ", _
    " BOENJINIERE ", " INGEGNERIA ", " ENGENHARIA ", " ENISINIA ", " INGINERIE ", " INNLEADAIREACHD ", " INGENIERISTIKO ", _
    " INGENIEURSWESE ", " INJINIA ", " REKAYASA ", " INGENIERIA ", " INGENIERIE ", " INGIGNIRIA ", " PROSJEKTERING ", _
    " BOUWKUNDE ", " AIKIN INJINIYA ", " DEIFBAU ")
    For i = 0 To UBound(theEngineer)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theEngineer(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theEngineer(i), " ENG ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theEngineer(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - AUTO
    theAutomation = Array(" AUTOMATION ", " AUTOMATISERING ", " OTOMATISASI ", " AUTOMATIZACION ", " AUTOMATIZACIJA ", _
    " AUTOMATISATION ", " OTOMASYON ", " AUTOMATIZAZIOA ", " AUTOMATIGO ", " AUTOMATITZACIO ", " OUTOMATISERING ", _
    " AUTOMATISEARRING ", " AUTOMATIKA ", " AKPAAKA ", " AUTOMATIZALAS ", " HAL NU NGAJADIKEUN OTOMATIS ", " AVTOMATLASDIRMA ", _
    " ZOKHA ", " AUTOMATYZACJA ", " AUTOMAATIO ", " AUTOMATIZIM ", " NGOKUZENZEKELAYO ", " UKUZENZEKELAYO ", " AWTOMEIDDIO ", _
    " AUTOMATIZAZIONE ", " OTOMATIKI ", " AUTOMASJON ", " AVTOMATLASHTIRISH ", " AWTOMATIZZAZZJONI ", " AUTOMASI ", " AUTOMATIQUE ", _
    " AUTOMATIZACE ", " AUTOMATIZACIA ", " AVTOMATIZACIJA ", " AUTOMATISEERIMINE ", " UATHOIBRIU ", " HO IKETSETSA LETHO ", _
    " AUTOMAZIONE ", " AUTOMACAO ", " AIKI DA KAI ", " AUTOMATISATIOUN ", " AUTOMATIZARE ", " FEIN-GHLUASAD ", _
    " AUTOMOTIVE ", " OTOMOTIV ", " AUTOMOBILGINTZA ", " AUTO ", " AUTOMOBIL ", " AUTOMOCIO ", " MOTOR ", " AUTOMOBILIAI ", _
    " AKURUNGWA ", " AUTOIPARI ", " OTOMOTIF ", " AUTOMOTOR ", " AUTOMOBILSKI ", " AVTOMOBIL ", " MAGALIMOTO ", " EGET ", _
    " AUTOMOBILU RUPNIECIBA ", " AUTOMOBILE ", " AUTOMOBILSKA ", " AUTOMOBILOWY ", " AUTOJEN ", " AUTOMOBILISTIK ", _
    " IMOTO ", " IZIMOTO ", " MODUROL ", " AUTOMOBILISTICA ", " TSHEB ", " OTOTOTIK ", " OTOMOBIL ", " BAABUURTA ", _
    " TAL-KAROZZI ", " AUTOMOTIF ", " FIARA ", " AUTOMOBILOVY PRUMYSL ", " MOTOKARI ", " MAGARI ", " AVTOMOBILSKA INDUSTRIJA ", _
    " AUTOTOOSTUS ", " FEITHICLEACH ", " BIL ", " LIKOLOI ", " SETTORE AUTOMOBILISTICO ", " AUTOMOTIVO ", " MOTA ", " TAAVALE AFI ", _
    " UIDHEAMACHD ")
    For i = 0 To UBound(theAutomation)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theAutomation(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theAutomation(i), " AUTO ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theAutomation(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - ENTERPRISE
    theEnterprise = Array(" ENTERPRISE ", " KURULUS ", " FORETAGENDE ", " ENPRESA ", " PERUSAHAAN ", " ENTREPRENO ", " EMPRESA ", _
    " ONDERNEMING ", " BEDRIUW ", " ?MONE ", " ULO ORU ", " VALLALKOZAS ", " PAUSAHAAN ", " PODUZECE ", " MUASSISA ", " MALONDA ", _
    " COEPTIS ", " UZNEMUMS ", " ENTREPRISE ", " PREDUZECE ", " PRZEDSIEBIORSTWO ", " YRITYS ", " NDERMARRJE ", " SHISHINI ", _
    " IBHIZINISI ", " MENTER ", " IMPRESE ", " SIRKET ", " BEDRIFTEN ", " ANTREPRIZ ", " KORXONA ", " SHIRKAD ", " INTRAPRIZA ", _
    " ORINASA ", " NEGOSYO ", " PODNIK ", " BUSINESS ", " BIASHARA ", " PODJETJE ", " ETTEVOTE ", " FIONTAR ", " FORETAG ", " KHOEBO ", _
    " IMPRESA ", " EMPREENDIMENTO ", " CINIKI ", " ATINAE ", " AFACERE ", " IOMAIRT ")
    For i = 0 To UBound(theEnterprise)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theEnterprise(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theEnterprise(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theEnterprise(i), 1)
    Wend
    Next i
    
End Function

Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("B1:B58") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("A1:A58") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:F58")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


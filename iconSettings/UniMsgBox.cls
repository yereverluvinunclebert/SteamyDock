VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* UniMsgBox class module
'* ----------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*
'* This class is a more comprehensive replacement of the standard MsgBox.
'* For a simpler one that easily replaces the original, see UnicodeMessageBox.bas
'*
'* CREDITS
'* -------
'* - Paul Caton and LaVolpe for their work on SelfSub, SelfHook and SelfCallback
'* - Randy Birch for the original example on changing MsgBox buttons
'*************************************************************************************************
Option Explicit

' ButtonText is only returned if it is customized
Public Event Result(ByVal Button As MsgBoxResult, ByVal ButtonText As String)

Public Enum MsgBoxButtons
    [OK Only] = vbOKOnly
    [OK / Cancel] = vbOKCancel
    [Abort / Retry / Ignore] = vbAbortRetryIgnore
    [Yes / No / Cancel] = vbYesNoCancel
    [Yes / No] = vbYesNo
    [Retry / Cancel] = vbRetryCancel
    [Cancel / Try / Continue] = &H6
End Enum

Public Enum MsgBoxDefault
    [Default Button 1] = vbDefaultButton1
    [Default Button 2] = vbDefaultButton2
    [Default Button 3] = vbDefaultButton3
    [Default Button 4] = vbDefaultButton4
End Enum

Public Enum MsgBoxIcon
    [Icon None] = &H0
    [Icon Critical] = &H10
    [Icon Question] = &H20
    [Icon Exclamation] = &H30
    [Icon Information] = &H40
    [Icon Resource] = &H80
End Enum

Public Enum MsgBoxLanguage
    [Invalid Language] = 0&
    [Language Afrikaans] = &H436  ' Afrikaans (South Africa)    af-ZA   Latn    1252
    [Language Albanian] = &H41C  ' Albanian (Albania)  sq-AL   Latn    1252
    [Language Alsatian] = &H484  ' Windows Vista and later: Alsatian (France)  gsw-FR
    [Language Amharic] = &H45E  ' Windows Vista and later: Amharic (Ethiopia)     am-ET       Unicode only
    [Language Arabic/Algeria] = &H1401 ' Arabic (Algeria)    ar-DZ   Arab    1256
    [Language Arabic/Bahrain] = &H3C01 ' Arabic (Bahrain)    ar-BH   Arab    1256
    [Language Arabic/Egypt] = &HC01  ' Arabic (Egypt)  ar-EG   Arab    1256
    [Language Arabic/Iraq] = &H801  ' Arabic (Iraq)   ar-IQ   Arab    1256
    [Language Arabic/Jordan] = &H2C01 ' Arabic (Jordan)     ar-JO   Arab    1256
    [Language Arabic/Kuwait] = &H3401 ' Arabic (Kuwait)     ar-KW   Arab    1256
    [Language Arabic/Lebanon] = &H3001 '' Arabic (Lebanon)    ar-LB   Arab    1256
    [Language Arabic/Libya] = &H1001 '' Arabic (Libya)  ar-LY   Arab    1256
    [Language Arabic/Morocco] = &H1801 '' Arabic (Morocco)    ar-MA   Arab    1256
    [Language Arabic/Oman] = &H2001 '' Arabic (Oman)   ar-OM   Arab    1256
    [Language Arabic/Qatar] = &H4001 '' Arabic (Qatar)  ar-QA   Arab    1256
    [Language Arabic/Saudia Arabia] = &H401  '' Arabic (Saudi Arabia)   ar-SA   Arab    1256
    [Language Arabic/Syria] = &H2801 '' Arabic (Syria)  ar-SY   Arab    1256
    [Language Arabic/Tunisia] = &H1C01 '' Arabic (Tunisia)    ar-TN   Arab    1256
    [Language Arabic/U.A.E.] = &H3801 '' Arabic (U.A.E.)     ar-AE   Arab    1256
    [Language Arabic/Yemen] = &H2401 '' Arabic (Yemen)  ar-YE   Arab    1256
    [Language Armenian] = &H42B  '' Windows 2000 and later: Armenian (Armenia)  hy-AM   Armn    Unicode only
    [Language Assamese] = &H44D  '' Windows Vista and later: Assamese (India)   as-IN       Unicode only
    [Language Azeri/Cyrillic] = &H82C  '' Azeri (Azerbaijan, Cyrillic)    az-Cyrl-AZ  Cyrl    1251
    [Language Azeri/Latin] = &H42C  '' Azeri (Azerbaijan, Latin)   az-Latn-AZ  Latn    1254
    [Language Bashkir] = &H46D  '' Windows Vista and later: Bashkir (Russia)   ba-RU
    [Language Basque] = &H42D  '' Basque (Basque)     eu-ES   Latn    1252
    [Language Belarusian] = &H423  '' Belarusian (Belarus)    be-BY   Cyrl    1251
    [Language Bengali] = &H445  '' Windows XP SP2 and later: Bengali (India)   bn-IN   Beng    Unicode only
    [Language Bosnian/Cyrillic] = &H201A '' Windows XP SP2 and later (downloadable); Windows Vista and later: Bosnian (Bosnia and Herzegovina, Cyrillic)    bs-Cyrl-BA  Cyrl    1251
    [Language Bosnian/Latin] = &H141A '' Windows XP SP2 and later: Bosnian (Bosnia and Herzegovina, Latin)   bs-Latn-BA  Latn    1250
    [Language Breton] = &H47E  '' Breton (France)     br-FR   Latn    1252
    [Language Bulgarian] = &H402  '' Bulgarian (Bulgaria)    bg-BG   Cyrl    1251
    [Language Burmese] = &H455  '' Not supported: Burmese
    [Language Catalan] = &H403  '' Catalan (Catalan)   ca-ES   Latn    1252
    [Language Chinese/Hong Kong SAR, PRC] = &HC04  '' Chinese (Hong Kong SAR, PRC)    zh-HK   Hant    950
    [Language Chinese/Macao SAR] = &H1404 '' Windows 98/Me, Windows XP and later: Chinese (Macao SAR)    zh-MO   Hant    950
    [Language Chinese/PRC] = &H804  '' Chinese (PRC)   zh-CN   Hans    936
    [Language Chinese/Singapore] = &H1004 '' Chinese (Singapore)     zh-SG   Hans    936
    [Language Chinese/Taiwan] = &H404  '' Chinese (Taiwan)    zh-TW   Hant    950
    [Language Croatian/Bosnia and Herzegovina/Latin] = &H101A '' Windows XP SP2 and later: Croatian (Bosnia and Herzegovina, Latin)  hr-BA   Latn    1250
    [Language Croatian] = &H41A  '' Croatian (Croatia)  hr-HR   Latn    1250
    [Language Czech] = &H405  '' Czech (Czech Republic)  cs-CZ   Latn    1250
    [Language Danish] = &H406  '' Danish (Denmark)    da-DK   Latn    1252
    [Language Dari] = &H48C  '' Windows XP and later: Dari (Afghanistan)    prs-AF  Arab    1256
    [Language Divehi] = &H465  '' Windows XP and later: Divehi (Maldives)     dv-MV   Thaa    Unicode only
    [Language Dutch/Belgium] = &H813  '' Dutch (Belgium)     nl-BE   Latn    1252
    [Language Dutch/Netherlands] = &H413  '' Dutch (Netherlands)     nl-NL   Latn    1252
    [Language English/Australia] = &HC09  '' English (Australia)     en-AU   Latn    1252
    [Language English/Belize] = &H2809 '' English (Belize)    en-BZ   Latn    1252
    [Language English/Canada] = &H1009 '' English (Canada)    en-CA   Latn    1252
    [Language English/Caribbean] = &H2409 '' English (Caribbean)     en-029  Latn    1252
    [Language English/India] = &H4009 '' Windows Vista and later: English (India)    en-IN   Latn    1252
    [Language English/Ireland] = &H1809 '' English (Ireland)   en-IE   Latn    1252
    [Language English/Jamaica] = &H2009 '' English (Jamaica)   en-JM   Latn    1252
    [Language English/Malaysia] = &H4409 '' Windows Vista and later: English (Malaysia)     en-MY   Latn    1252
    [Language English/New Zealand] = &H1409 '' English (New Zealand)   en-NZ   Latn    1252
    [Language English/Philippines] = &H3409 '' Windows 98/Me, Windows 2000 and later: English (Philippines)    en-PH   Latn    1252
    [Language English/Singapore] = &H4809 '' Windows Vista and later: English (Singapore)    en-SG   Latn    1252
    [Language English/South Africa] = &H1C09 '' English (South Africa)  en-ZA   Latn    1252
    [Language English/Trinidad and Tobago] = &H2C09 '' English (Trinidad and Tobago)   en-TT   Latn    1252
    [Language English/United Kingdom] = &H809  '' English (United Kingdom)    en-GB   Latn    1252
    [Language English/United States] = &H409  '' English (United States)     en-US   Latn    1252
    [Language English/Zimbabwe] = &H3009 '' Windows 98/Me, Windows 2000 and later: English (Zimbabwe)   en-ZW   Latn    1252
    [Language Estonian] = &H425  '' Estonian (Estonia)  et-EE   Latn    1257
    [Language Faroese] = &H438  '' Faroese (Faroe Islands)     fo-FO   Latn    1252
    [Language Filipino] = &H464  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Filipino (Philippines)    fil-PH  Latn    1252
    [Language Finnish] = &H40B  '' Finnish (Finland)   fi-FI   Latn    1252
    [Language French/Belgium] = &H80C  '' French (Belgium)    fr-BE   Latn    1252
    [Language French/Canada] = &HC0C  '' French (Canada)     fr-CA   Latn    1252
    [Language French/France] = &H40C  '' French (France)     fr-FR   Latn    1252
    [Language French/Luxembourg] = &H140C '' French (Luxembourg)     fr-LU   Latn    1252
    [Language French/Monaco] = &H180C '' French (Monaco)     fr-MC   Latn    1252
    [Language French/Switzerland] = &H100C '' French (Switzerland)    fr-CH   Latn    1252
    [Language Frisian] = &H462  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Frisian (Netherlands)     fy-NL   Latn    1252
    [Language Galician] = &H456  '' Windows XP and later: Galician (Spain)  gl-ES   Latn    1252
    [Language Georgian] = &H437  '' Windows 2000 and later: Georgian (Georgia)  ka-GE   Geor    Unicode only
    [Language German/Austria] = &HC07  '' German (Austria)    de-AT   Latn    1252
    [Language German/Germany] = &H407  '' German (Germany)    de-DE   Latn    1252
    [Language German/Liechtenstein] = &H1407 '' German (Liechtenstein)  de-LI   Latn    1252
    [Language German/Luxembourg] = &H1007 '' German (Luxembourg)     de-LU   Latn    1252
    [Language German/Switzerland] = &H807  '' German (Switzerland)    de-CH   Latn    1252
    [Language Greek] = &H408  '' Greek (Greece)  el-GR   Grek    1253
    [Language Greenlandic] = &H46F  '' Windows Vista and later: Greenlandic (Greenland)    kl-GL   Latn    1252
    [Language Gujarati] = &H447  '' Windows XP and later: Gujarati (India)  gu-IN   Gujr    Unicode only
    [Language Hausa] = &H468  '' Windows Vista and later: Hausa (Nigeria, Latin)     ha-Latn-NG  Latn    1252
    [Language Hebrew] = &H40D  '' Hebrew (Israel)     he-IL   Hebr    1255
    [Language Hindi] = &H439  '' Windows 2000 and later: Hindi (India)   hi-IN   Deva    Unicode only
    [Language Hungarian] = &H40E  '' Hungarian (Hungary)     hu-HU   Latn    1250
    [Language Icelandic] = &H40F  '' Icelandic (Iceland)     is-IS   Latn    1252
    [Language Igbo] = &H470  '' Igbo (Nigeria)  ig-NG
    [Language Indonesian] = &H421  '' Indonesian (Indonesia)  id-ID   Latn    1252
    [Language Inuktitut/Latin] = &H85D  '' Windows XP and later: Inuktitut (Canada, Latin)     iu-Latn-CA  Latn    1252
    [Language Inuktitut/Syllabics] = &H45D  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Inuktitut (Canada, Syllabics)     iu-Cans-CA  Cans    Unicode only
    [Language Irish] = &H83C  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Irish (Ireland)   ga-IE   Latn    1252
    [Language Italian/Italy] = &H410  '' Italian (Italy)     it-IT   Latn    1252
    [Language Italian/Switzerland] = &H810  '' Italian (Switzerland)   it-CH   Latn    1252
    [Language Japanese] = &H411  '' Japanese (Japan)    ja-JP   Hani;Hira;Kana  932
    [Language Kannada] = &H44B  '' Windows XP and later: Kannada (India)   kn-IN   Knda    Unicode only
    [Language Kazakh] = &H43F  '' Windows 2000 and later: Kazakh (Kazakhstan)     kk-KZ   Cyrl    1251
    [Language Khmer] = &H453  '' Windows Vista and later: Khmer (Cambodia)   kh-KH   Khmr    Unicode only
    [Language K'iche] = &H486  '' Windows Vista And later: K 'iche (Guatemala)     qut-GT  Latn    1252
    [Language Kinyarwanda] = &H487  '' Windows Vista and later: Kinyarwanda (Rwanda)   rw-RW   Latn    1252
    [Language Konkani] = &H457  '' Windows 2000 and later: Konkani (India)     kok-IN  Deva    Unicode only
    [Language Korean/Johab] = &H812  '' Windows 95, Windows NT 4.0 only: Korean (Johab)
    [Language Korean/Korea] = &H412  '' Korean (Korea)  ko-KR   Hang;Hani   949
    [Language Kyrgyz] = &H440  '' Windows XP and later: Kyrgyz (Kyrgyzstan)   ky-KG   Cyrl    1251
    [Language Lao] = &H454  '' Windows Vista and later: Lao (Lao PDR)  lo-LA   Laoo    Unicode only
    [Language Latvian] = &H426  '' Latvian (Latvia)    lv-LV   Latn    1257
    [Language Lithuanian] = &H427  '' Lithuanian (Lithuania)  lt-LT   Latn    1257
    [Language Lower Sorbian] = &H82E  '' Windows Vista and later: Lower Sorbian (Germany)    dsb-DE  Latn    1252
    [Language Luxembourgish] = &H46E  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Luxembourgish (Luxembourg)    lb-LU   Latn    1252
    [Language Macedonia] = &H42F  '' Windows 2000 and later: Macedonian (Macedonia, FYROM)   mk-MK   Cyrl    1251
    [Language Malay/Brunei Darussalam] = &H83E  '' Windows 2000 and later: Malay (Brunei Darussalam)   ms-BN   Latn    1252
    [Language Malay/Malaysia] = &H43E  '' Windows 2000 and later: Malay (Malaysia)    ms-MY   Latn    1252
    [Language Malayalam] = &H44C  '' Windows XP SP2 and later: Malayalam (India)     ml-IN   Mlym    Unicode only
    [Language Maltese] = &H43A  '' Windows XP SP2 and later: Maltese (Malta)   mt-MT   Latn    1252
    [Language Maori] = &H481  '' Windows XP SP2 and later: Maori (New Zealand)   mi-NZ   Latn    1252
    [Language Mapudungun] = &H47A  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Mapudungun (Chile)    arn-CL  Latn    1252
    [Language Marathi] = &H44E  '' Windows 2000 and later: Marathi (India)     mr-IN   Deva    Unicode only
    [Language Mohawk] = &H47C  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Mohawk (Canada)   moh-CA  Latn    1252
    [Language Mongolian/Mongolia] = &H450  '' Windows XP and later: Mongolian (Mongolia)  mn-Cyrl-MN  Cyrl    1251
    [Language Mongolian/PRC] = &H850  '' Windows Vista and later: Mongolian (PRC)    mn-Mong-CN  Mong    Unicode only
    [Language Nepali] = &H461  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Nepali (Nepal)    ne-NP   Deva    Unicode only
    [Language Norwegian/Bokmål] = &H414  '' Norwegian (Bokmål, Norway)  nb-NO   Latn    1252
    [Language Norwegian/Nynorsk] = &H814  '' Norwegian (Nynorsk, Norway)     nn-NO   Latn    1252
    [Language Occitan] = &H482  '' Occitan (France)    oc-FR   Latn    1252
    [Language Oriya] = &H448  '' Oriya (India)   or-IN   Orya    Unicode only
    [Language Pashto] = &H463  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Pashto (Afghanistan)  ps-AF
    [Language Persian] = &H429  '' Persian (Iran)  fa-IR   Arab    1256
    [Language Polish] = &H415  '' Polish (Poland)     pl-PL   Latn    1250
    [Language Portuguese/Brazil] = &H416  '' Portuguese (Brazil)     pt-BR   Latn    1252
    [Language Portuguese/Portugal] = &H816  '' Portuguese (Portugal)   pt-PT   Latn    1252
    [Language Punjabi] = &H446  '' Windows XP and later: Punjabi (India)   pa-IN   Guru    Unicode only
    [Language Quechua/Bolivia] = &H46B  '' Windows XP SP2 and later: Quechua (Bolivia)     quz-BO  Latn    1252
    [Language Quechua/Ecuador] = &H86B  '' Windows XP SP2 and later: Quechua (Ecuador)     quz-EC  Latn    1252
    [Language Quechua/Peru] = &HC6B  '' Windows XP SP2 and later: Quechua (Peru)    quz-PE  Latn    1252
    [Language Romanian] = &H418  '' Romanian (Romania)  ro-RO   Latn    1250
    [Language Romansh] = &H417  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Romansh (Switzerland)     rm-CH   Latn    1252
    [Language Russian] = &H419  '' Russian (Russia)    ru-RU   Cyrl    1251
    [Language Sami/Inari] = &H243B '' Windows XP SP2 and later: Sami (Inari, Finland)     smn-FI  Latn    1252
    [Language Sami/Lule/Norway] = &H103B '' Windows XP SP2 and later: Sami (Lule, Norway)   smj-NO  Latn    1252
    [Language Sami/Lule/Sweden] = &H143B '' Windows XP SP2 and later: Sami (Lule, Sweden)   smj-SE  Latn    1252
    [Language Sami/Northern/Finland] = &HC3B  '' Windows XP SP2 and later: Sami (Northern, Finland)  se-FI   Latn    1252
    [Language Sami/Northern/Norway] = &H43B  '' Windows XP SP2 and later: Sami (Northern, Norway)   se-NO   Latn    1252
    [Language Sami/Northern/Sweden] = &H83B  '' Windows XP SP2 and later: Sami (Northern, Sweden)   se-SE   Latn    1252
    [Language Sami/Skolt/Finland] = &H203B '' Windows XP SP2 and later: Sami (Skolt, Finland)     sms-FI  Latn    1252
    [Language Sami/Southern/Norway] = &H183B '' Windows XP SP2 and later: Sami (Southern, Norway)   sma-NO  Latn    1252
    [Language Sami/Southern/Sweden] = &H1C3B '' Windows XP SP2 and later: Sami (Southern, Sweden)   sma-SE  Latn    1252
    [Language Sanskrit] = &H44F  '' Windows 2000 and later: Sanskrit (India)    sa-IN   Deva    Unicode only
    [Language Serbian/Bosnia and Herzegovina/Cyrillic] = &H1C1A '' Windows XP SP2 and later: Serbian (Bosnia and Herzegovina, Cyrillic)    sr-Cyrl-BA  Cyrl    1251
    [Language Serbian/Bosnia and Herzegovina/Latin] = &H181A '' Windows XP SP2 and later: Serbian (Bosnia and Herzegovina, Latin)   sr-Latn-BA  Latn    1250
    [Language Serbian/Serbia/Cyrillic] = &HC1A  '' Serbian (Serbia, Cyrillic)  sr-Cyrl-CS  Cyrl    1251
    [Language Serbian/Serbia/Latin] = &H81A  '' Serbian (Serbia, Latin)     sr-Latn-CS  Latn    1250
    [Language Sesotho sa Leboa] = &H46C  '' Windows XP SP2 and later: Sesotho sa Leboa/Northern Sotho (South Africa)    ns-ZA   Latn    1252
    [Language Setswana / Tswana] = &H432  '' Windows XP SP2 and later: Setswana/Tswana (South Africa)    tn-ZA   Latn    1252
    [Language Sinhala] = &H45B  '' Windows Vista and later: Sinhala (Sri Lanka)    si-LK   Sinh    Unicode only
    [Language Slovak] = &H41B  '' Slovak (Slovakia)   sk-SK   Latn    1250
    [Language Slovenian] = &H424  '' Slovenian (Slovenia)    sl-SI   Latn    1250
    [Language Spanish/Argentina] = &H2C0A '' Spanish (Argentina)     es-AR   Latn    1252
    [Language Spanish/Bolivia] = &H400A '' Spanish (Bolivia)   es-BO   Latn    1252
    [Language Spanish/Chile] = &H340A '' Spanish (Chile)     es-CL   Latn    1252
    [Language Spanish/Colombia] = &H240A '' Spanish (Colombia)  es-CO   Latn    1252
    [Language Spanish/Costa Rica] = &H140A '' Spanish (Costa Rica)    es-CR   Latn    1252
    [Language Spanish/Dominican Republic] = &H1C0A '' Spanish (Dominican Republic)    es-DO   Latn    1252
    [Language Spanish/Ecuador] = &H300A '' Spanish (Ecuador)   es-EC   Latn    1252
    [Language Spanish/El Salvador] = &H440A '' Spanish (El Salvador)   es-SV   Latn    1252
    [Language Spanish/Guatemala] = &H100A '' Spanish (Guatemala)     es-GT   Latn    1252
    [Language Spanish/Honduras] = &H480A '' Spanish (Honduras)  es-HN   Latn    1252
    [Language Spanish/Mexico] = &H80A  '' Spanish (Mexico)    es-MX   Latn    1252
    [Language Spanish/Nicaragua] = &H4C0A '' Spanish (Nicaragua)     es-NI   Latn    1252
    [Language Spanish/Panama] = &H180A '' Spanish (Panama)    es-PA   Latn    1252
    [Language Spanish/Paraguay] = &H3C0A '' Spanish (Paraguay)  es-PY   Latn    1252
    [Language Spanish/Peru] = &H280A '' Spanish (Peru)  es-PE   Latn    1252
    [Language Spanish/Puerto Rico] = &H500A '' Spanish (Puerto Rico)   es-PR   Latn    1252
    [Language Spanish/Spain] = &HC0A  '' Spanish (Spain)     es-ES   Latn    1252
    [Language Spanish/Spain Traditional] = &H40A  '' Spanish (Spain, Traditional Sort)   es-ES_tradnl    Latn    1252
    [Language Spanish/United States] = &H540A '' Windows Vista and later: Spanish (United States)    es-US
    [Language Spanish/Uruguay] = &H380A '' Spanish (Uruguay)   es-UY   Latn    1252
    [Language Spanish/Venezuela] = &H200A '' Spanish (Venezuela)     es-VE   Latn    1252
    [Language Sutu] = &H430  '' Not supported: Sutu
    [Language Swahili] = &H441  '' Windows 2000 and later: Swahili (Kenya)     sw-KE   Latn    1252
    [Language Swedish/Finland] = &H81D  '' Swedish (Finland)   sv-FI   Latn    1252
    [Language Swedish/Sweden] = &H41D  '' Swedish (Sweden)    sv-SE   Latn    1252
    [Language Syriac] = &H45A  '' Windows XP and later: Syriac (Syria)    syr-SY  Syrc    Unicode only
    [Language Tajik] = &H428  '' Windows Vista and later: Tajik (Tajikistan)     tg-Cyrl-TJ  Cyrl    1251
    [Language Tamazight] = &H85F  '' Windows Vista and later: Tamazight (Algeria, Latin)     tmz-Latn-DZ     Latn    1252
    [Language Tamil] = &H449  '' Windows 2000 and later: Tamil (India)   ta-IN   Taml    Unicode only
    [Language Tatar] = &H444  '' Windows XP and later: Tatar (Russia)    tt-RU   Cyrl    1251
    [Language Telugu] = &H44A  '' Windows XP and later: Telugu (India)    te-IN   Telu    Unicode only
    [Language Thai] = &H41E  '' Thai (Thailand)     th-TH   Thai    874
    [Language Tibetan/Bhutan] = &H851  '' Windows Vista and later: Tibetan (Bhutan)   bo-BT   Tibt    Unicode only
    [Language Tibetan/PRC] = &H451  '' Windows Vista and later: Tibetan (PRC)  bo-CN   Tibt    Unicode only
    [Language Turkish] = &H41F  '' Turkish (Turkey)    tr-TR   Latn    1254
    [Language Turkmen] = &H442  '' Windows Vista and later: Turkmen (Turkmenistan)     tk-TM   Cyrl    1251
    [Language Uighur] = &H480  '' Windows Vista and later: Uighur (PRC)   ug-CN   Arab    1256
    [Language Ukrainian] = &H422  '' Ukrainian (Ukraine)     uk-UA   Cyrl    1251
    [Language Upper Sorbian] = &H42E  '' Windows Vista and later: Upper Sorbian (Germany)    wen-DE  Latn    1252
    [Language Urdu/India] = &H820  '' Urdu (India)    tr-IN
    [Language Urdu/Pakistan] = &H420  '' Windows 98/Me, Windows 2000 and later: Urdu (Pakistan)  ur-PK   Arab    1256
    [Language Uzbek/Cyrillic] = &H843  '' Windows 2000 and later: Uzbek (Uzbekistan, Cyrillic)    uz-Cyrl-UZ  Cyrl    1251
    [Language Uzbek/Latin] = &H443  '' Windows 2000 and later: Uzbek (Uzbekistan, Latin)   uz-Latn-UZ  Latn    1254
    [Language Vietnamese] = &H42A  '' Windows 98/Me, Windows NT 4.0 and later: Vietnamese (Vietnam)   vi-VN   Latn    1258
    [Language Welsh] = &H452  '' Windows XP SP2 and later: Welsh (United Kingdom)    cy-GB   Latn    1252
    [Language Wolof] = &H488  '' Windows Vista and later: Wolof (Senegal)    wo-SN   Latn    1252
    [Language Xhosa / isiXhosa] = &H434  '' Windows XP SP2 and later: Xhosa/isiXhosa (South Africa)     xh-ZA   Latn    1252
    [Language Yakut] = &H485  '' Windows Vista and later: Yakut (Russia)     sah-RU  Cyrl    1251
    [Language Yi] = &H478  '' Windows Vista and later: Yi (PRC)   ii-CN   Yiii    Unicode only
    [Language Yoruba] = &H46A  '' Windows Vista and later: Yoruba (Nigeria)   yo-NG
    [Language Zulu / isiZulu] = &H435  '' Windows XP SP2 and later: Zulu/isiZulu (South Africa)   zu-ZA   Latn    1252
End Enum

Public Enum MsgBoxModality
    [MsgBox Modeless] = -1
    [MsgBox Application] = 0
    [MsgBox System] = &H1000
    [MsgBox Task] = &H2000
End Enum

Public Enum MsgBoxResult
    [Result OK] = vbOK
    [Result Cancel] = vbCancel
    [Result Abort] = vbAbort
    [Result Retry] = vbRetry
    [Result Ignore] = vbIgnore
    [Result Yes] = vbYes
    [Result No] = vbNo
    [Result Try Again] = &HA
    [Result Continue] = &HB
End Enum

Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5
Private Const WM_SETICON = &H80

Private Declare Function MessageBoxIndirectW Lib "user32" (lpMsgBoxParams As MsgBoxParams) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetDlgItemTextW Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Type MsgBoxParams
    cbSize As Long
    hWndOwner As Long
    hInstance As Long
    lpszText As Long
    lpszCaption As Long
    dwStyle As Long
    lpszIcon As Long
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Private m_ButtonAbort As String
Private m_ButtonCancel As String
Private m_ButtonContinue As String
Private m_ButtonIgnore As String
Private m_ButtonNo As String
Private m_ButtonOK As String
Private m_ButtonRetry As String
Private m_ButtonTry As String
Private m_ButtonYes As String
Private m_hWnd As Long
Private m_Icon As Long
Private m_Modality As MsgBoxModality
Private m_Params As MsgBoxParams
Private m_Resource As String
Private m_Title As String

Private m_Hook As Long
Private m_HookProc As Long

' the following are the requirements for SelfCallback
Private Const IDX_CALLBACKORDINAL As Long = 36

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Sub Class_Initialize()
    ' replicate native MsgBox settings
    m_Modality = [MsgBox Application]
    m_Title = App.Title
    With m_Params
        .cbSize = Len(m_Params)
        .hInstance = App.hInstance
    End With
End Sub
Private Sub Class_Terminate()
    ' just in case
    RemoveHook
End Sub
Public Property Get AlignRight() As Boolean
    AlignRight = (m_Params.dwStyle And vbMsgBoxRight) = vbMsgBoxRight
End Property
Public Property Let AlignRight(ByVal NewValue As Boolean)
    If NewValue Then
        m_Params.dwStyle = m_Params.dwStyle Or vbMsgBoxRight
    Else
        m_Params.dwStyle = m_Params.dwStyle And Not vbMsgBoxRight
    End If
End Property
Public Property Get AlwaysOnTop() As Boolean
    AlwaysOnTop = (m_Params.dwStyle And &H40000) = &H40000
End Property
Public Property Let AlwaysOnTop(ByVal NewValue As Boolean)
    If NewValue Then
        m_Params.dwStyle = m_Params.dwStyle Or &H40000
    Else
        m_Params.dwStyle = m_Params.dwStyle And Not &H40000
    End If
End Property
Public Property Get ButtonAbort() As String
    ButtonAbort = m_ButtonAbort
End Property
Public Property Let ButtonAbort(ByRef NewValue As String)
    m_ButtonAbort = NewValue
End Property
Public Property Get ButtonCancel() As String
    ButtonCancel = m_ButtonCancel
End Property
Public Property Let ButtonCancel(ByRef NewValue As String)
    m_ButtonCancel = NewValue
End Property
Public Property Get ButtonContinue() As String
    ButtonContinue = m_ButtonContinue
End Property
Public Property Let ButtonContinue(ByRef NewValue As String)
    m_ButtonContinue = NewValue
End Property
Public Property Get ButtonIgnore() As String
    ButtonIgnore = m_ButtonIgnore
End Property
Public Property Let ButtonIgnore(ByRef NewValue As String)
    m_ButtonIgnore = NewValue
End Property
Public Property Get ButtonNo() As String
    ButtonNo = m_ButtonNo
End Property
Public Property Let ButtonNo(ByRef NewValue As String)
    m_ButtonNo = NewValue
End Property
Public Property Get ButtonOK() As String
    ButtonOK = m_ButtonOK
End Property
Public Property Let ButtonOK(ByRef NewValue As String)
    m_ButtonOK = NewValue
End Property
Public Property Get ButtonRetry() As String
    ButtonRetry = m_ButtonRetry
End Property
Public Property Let ButtonRetry(ByRef NewValue As String)
    m_ButtonRetry = NewValue
End Property
Public Property Get ButtonTryAgain() As String
    ButtonTryAgain = m_ButtonTry
End Property
Public Property Let ButtonTryAgain(ByRef NewValue As String)
    m_ButtonTry = NewValue
End Property
Public Property Get ButtonYes() As String
    ButtonYes = m_ButtonYes
End Property
Public Property Let ButtonYes(ByRef NewValue As String)
    m_ButtonYes = NewValue
End Property
Public Property Get Caption() As String
    Caption = m_Title
End Property
Public Property Let Caption(ByRef NewValue As String)
    m_Title = NewValue
End Property
Public Property Get Foreground() As Boolean
    Foreground = (m_Params.dwStyle And vbMsgBoxSetForeground) = vbMsgBoxSetForeground
End Property
Public Property Let Foreground(ByVal NewValue As Boolean)
    If NewValue Then
        m_Params.dwStyle = m_Params.dwStyle Or vbMsgBoxSetForeground
    Else
        m_Params.dwStyle = m_Params.dwStyle And Not vbMsgBoxSetForeground
    End If
End Property
Public Property Get HelpButton() As Boolean
    HelpButton = (m_Params.dwStyle And vbMsgBoxHelpButton) = vbMsgBoxHelpButton
End Property
Public Property Let HelpButton(ByVal NewValue As Boolean)
    If NewValue Then
        m_Params.dwStyle = m_Params.dwStyle Or vbMsgBoxHelpButton
    Else
        m_Params.dwStyle = m_Params.dwStyle And Not vbMsgBoxHelpButton
    End If
End Property
Public Property Get HelpContext() As Long
    HelpContext = m_Params.dwContextHelpId
End Property
Public Property Let HelpContext(ByVal NewValue As Long)
    m_Params.dwContextHelpId = NewValue
End Property
Public Property Get Icon() As Long
    Icon = m_Icon
End Property
Public Property Let Icon(ByVal NewValue As Long)
    m_Icon = NewValue
End Property
Public Property Get hWndOwner() As Long
    hWndOwner = m_hWnd
End Property
Public Property Let hWndOwner(ByVal NewValue As Long)
    m_hWnd = NewValue
End Property
Public Property Get Language() As MsgBoxLanguage
    Language = m_Params.dwLanguageId
End Property
Public Property Let Language(ByVal NewValue As MsgBoxLanguage)
    m_Params.dwLanguageId = NewValue And &HFFFF&
End Property
Public Property Get Modality() As MsgBoxModality
    Modality = m_Modality
End Property
Public Property Let Modality(ByVal NewValue As MsgBoxModality)
    m_Modality = NewValue
End Property
Private Sub RemoveHook()
    Const MEM_RELEASE As Long = &H8000&
    ' remove the hook
    If m_Hook Then UnhookWindowsHookEx m_Hook: m_Hook = 0
    ' only do this if we still have the procedure
    If m_HookProc Then
        ' free the procedure callback
        VirtualFree m_HookProc, 0, MEM_RELEASE
        ' reset procedure pointer to prevent this getting ran twice
        m_HookProc = 0
    End If
End Sub
Public Property Get ResourceIcon() As String
    ResourceIcon = m_Resource
End Property
Public Property Let ResourceIcon(ByRef NewValue As String)
    m_Resource = NewValue
    m_Params.lpszIcon = StrPtr(m_Resource)
End Property
Public Property Get RightToLeft() As Boolean
    RightToLeft = (m_Params.dwStyle And vbMsgBoxRtlReading) = vbMsgBoxRtlReading
End Property
Public Property Let RightToLeft(ByVal NewValue As Boolean)
    If NewValue Then
        m_Params.dwStyle = m_Params.dwStyle Or vbMsgBoxRtlReading
    Else
        m_Params.dwStyle = m_Params.dwStyle And Not vbMsgBoxRtlReading
    End If
End Property
Private Function ScAddressOfOrdinal(ByVal Ordinal As Long) As Long
    Dim bytValue As Byte, bytSignature As Byte, lngA As Long, lngAddress As Long
    ' get address of this class module instance
    RtlMoveMemory VarPtr(lngAddress), ObjPtr(Me), 4&
    ' probe for Class method
    If ScProbe(lngAddress + &H1C, lngA, bytSignature) Then
        ' scan up to 256 vTable entries
        For lngA = lngA + 4 To lngA + 1024 Step 4
            ' get address in vTable
            RtlMoveMemory VarPtr(lngAddress), lngA, 4&
            ' invalid code address?
            If IsBadCodePtr(lngAddress) Then
                ' return this vTable address
                RtlMoveMemory VarPtr(ScAddressOfOrdinal), lngA - (Ordinal * 4&), 4&
                Exit For
            End If
            ' get byte pointed to vTable address
            RtlMoveMemory VarPtr(bytValue), lngAddress, 1&
            ' if does not match the expected value...
            If bytValue <> bytSignature Then
                ' return this vTable address
                RtlMoveMemory VarPtr(ScAddressOfOrdinal), lngA - (Ordinal * 4&), 4&
                Exit For
            End If
        Next lngA
    End If
End Function
Private Property Get ScData(ByVal Index As Long, ByVal ThunkPtr As Long) As Long
    RtlMoveMemory VarPtr(ScData), ThunkPtr + (Index * 4&), 4&
End Property
Private Property Let ScData(ByVal Index As Long, ByVal ThunkPtr As Long, ByVal NewValue As Long)
    RtlMoveMemory ThunkPtr + (Index * 4&), VarPtr(NewValue), 4&
End Property
Private Function ScProbe(ByVal Address As Long, ByRef Method As Long, ByRef Signature As Byte) As Boolean
    Dim bytValue As Byte, lngVTableEntry As Long
    ' probe eight entries
    For Address = Address To Address + 32 Step 4
        ' get vTable entry
        RtlMoveMemory VarPtr(lngVTableEntry), Address, 4&
        ' if not an implemented interface
        If lngVTableEntry Then
            ' get the value pointed at by the vTable entry
            RtlMoveMemory VarPtr(bytValue), lngVTableEntry, 1&
            ' if native or P-code signature...
            If (bytValue = &H33) Or (bytValue = &HE9) Then
                ' return this information
                Method = Address
                Signature = bytValue
                ' success, exit loop
                ScProbe = True
                Exit For
            End If
        End If
    Next Address
End Function
Private Function ScProcedureAddress(ByVal DynamicLinkLibrary As String, ByVal Procedure As String, ByVal Unicode As Boolean) As Long
    ' get the procedure address
    If Unicode Then
        ScProcedureAddress = GetProcAddress(GetModuleHandleW(StrPtr(DynamicLinkLibrary)), Procedure)
    Else
        ScProcedureAddress = GetProcAddress(GetModuleHandleA(DynamicLinkLibrary), Procedure)
    End If
    ' in IDE, verify we got it
    Debug.Assert ScProcedureAddress
End Function
Public Sub SetButtons(Optional ByRef Abort As String, Optional ByRef Cancel As String, Optional ByRef Continue As String, Optional ByRef Ignore As String, Optional ByRef No As String, Optional ByRef OK As String, Optional ByRef Retry As String, Optional ByRef TryAgain As String, Optional ByRef Yes As String)
    m_ButtonAbort = Abort
    m_ButtonCancel = Cancel
    m_ButtonContinue = Continue
    m_ButtonIgnore = Ignore
    m_ButtonNo = No
    m_ButtonOK = OK
    m_ButtonRetry = Retry
    m_ButtonTry = TryAgain
    m_ButtonYes = Yes
End Sub
Private Sub SetHook()
    Dim lngCallback As Long, lngCallbackCode() As Long, lngScMem As Long

    ' index of ordinal (always 1 in this class, the last procedure)
    Const Ordinal As Long = 1
    ' number of parameters (for WH_CBT hook procedure this is always 3, thus this is a constant)
    Const ParamCount As Long = 3

    ' memory bytes required for the callback thunk
    Const MEM_LEN As Long = IDX_CALLBACKORDINAL * 4 + 4
    ' allocate executable memory
    Const PAGE_RWX As Long = &H40&
    ' commit allocated memory
    Const MEM_COMMIT As Long = &H1000&
    ' name of this routine
    Const SUB_NAME As String = "TimerSet"
    ' thunk data index of the Owner object's vTable address
    Const INDX_OWNER As Long = 0
    ' thunk data index of the callback procedure address
    Const INDX_CALLBACK As Long = 1
    ' thunk data index of the EbMode function address
    Const INDX_EBMODE As Long = 2
    ' thunk data index of the IsBadCodePtr function address
    Const INDX_BADPTR As Long = 3
    ' thunk data index of the KillTimer function address
    Const INDX_KT As Long = 4
    ' thunk code patch index of the thunk data
    Const INDX_EBX As Long = 6
    ' thunk code patch index of the number of parameters expected in callback
    Const INDX_PARAMS As Long = 18
    ' thunk code patch index of the bytes to be released after callback
    Const INDX_PARAMLEN As Long = 24
    ' thunk offset to the callback execution address
    Const PROC_OFF As Long = &H14

    ' start only if we have not done this already
    If m_HookProc = 0 Then
        ' get address of the hook procedure, the last procedure in this class module (by ordinal)
        lngCallback = ScAddressOfOrdinal(Ordinal)
        ' verify we got it
        If lngCallback Then
            ' allocate executable memory
            lngScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)
            ' verify we got it
            If lngScMem Then
                ' allocate for the machine-code array
                ReDim lngCallbackCode(0 To IDX_CALLBACKORDINAL) As Long
                ' create the machine-code array
                lngCallbackCode(5) = &HBB60E089
                lngCallbackCode(7) = &H73FFC589
                lngCallbackCode(8) = &HC53FF04
                lngCallbackCode(9) = &H59E80A74
                lngCallbackCode(10) = &HE9000000
                lngCallbackCode(11) = &H30&
                lngCallbackCode(12) = &H87B81
                lngCallbackCode(13) = &H75000000
                lngCallbackCode(14) = &H9090902B
                lngCallbackCode(15) = &H42DE889
                lngCallbackCode(16) = &H50000000
                lngCallbackCode(17) = &HB9909090
                lngCallbackCode(19) = &H90900AE3
                lngCallbackCode(20) = &H8D74FF
                lngCallbackCode(21) = &H9090FAE2
                lngCallbackCode(22) = &H53FF33FF
                lngCallbackCode(23) = &H90909004
                lngCallbackCode(24) = &H2BADC261
                lngCallbackCode(25) = &H3D0853FF
                lngCallbackCode(26) = &H1&
                lngCallbackCode(27) = &H23DCE74
                lngCallbackCode(28) = &H74000000
                lngCallbackCode(29) = &HAE807
                lngCallbackCode(30) = &H90900000
                lngCallbackCode(31) = &H4589C031
                lngCallbackCode(32) = &H90DDEBFC
                lngCallbackCode(33) = &HFF0C75FF
                lngCallbackCode(34) = &H53FF0475
                lngCallbackCode(35) = &HC310&
                ' settings
                lngCallbackCode(INDX_BADPTR) = ScProcedureAddress("kernel32", "IsBadCodePtr", False)
                ' set owner (this class)
                lngCallbackCode(INDX_OWNER) = ObjPtr(Me)
                ' callback address (the last procedure in this class)
                lngCallbackCode(INDX_CALLBACK) = lngCallback
                ' remember the ordinal used
                lngCallbackCode(IDX_CALLBACKORDINAL) = Ordinal
                ' parameter count
                lngCallbackCode(INDX_PARAMS) = ParamCount
                RtlMoveMemory VarPtr(lngCallbackCode(INDX_PARAMLEN)) + 2&, VarPtr(ParamCount * 4&), 2&
                ' IDE safety
                If App.LogMode = 0 Then
                    ' store the EbMode function address in the thunk data
                    lngCallbackCode(INDX_EBMODE) = ScProcedureAddress("vba6", "EbMode", False)
                End If
                ' special for timer callback:
                'lngCallbackCode(INDX_KT) = ScProcedureAddress("user32", "KillTimer", False)
                ' set the data address relative to virtual memory pointer
                lngCallbackCode(INDX_EBX) = lngScMem
                ' copy thunk code to executable memory
                RtlMoveMemory lngScMem, VarPtr(lngCallbackCode(INDX_OWNER)), MEM_LEN
                ' remember the procedure address
                m_HookProc = lngScMem + PROC_OFF
                ' now we can initialize the hook
                m_Hook = SetWindowsHookExW(WH_CBT, m_HookProc, App.hInstance, App.ThreadID)
            End If
        End If
    End If
End Sub
Public Function Show(ByVal Message As String, Optional Caption As String, Optional ByVal Buttons As MsgBoxButtons, Optional ByVal DefaultButton As MsgBoxDefault, Optional ByVal MessageIcon As MsgBoxIcon, Optional ByVal WindowIcon As Long) As MsgBoxResult
    Dim blnHook As Boolean, lngIcon As Long
    ' determine owner of the message box
    If m_Modality = [MsgBox Modeless] Then
        ' not a modal message box, thus no owner
        m_Params.hWndOwner = 0
    ElseIf m_hWnd = 0 Then
        ' check for active form
        If Not Screen.ActiveForm Is Nothing Then
            m_Params.hWndOwner = Screen.ActiveForm.hwnd
        Else
            ' forced to modeless
            m_Params.hWndOwner = 0
        End If
    Else
        ' use the given hWnd
        m_Params.hWndOwner = m_hWnd
    End If
    m_Params.lpszText = StrPtr(Message)
    If LenB(Caption) Then m_Params.lpszCaption = StrPtr(Caption) Else m_Params.lpszCaption = StrPtr(m_Title)
    ' set icon style
    m_Params.dwStyle = m_Params.dwStyle Or Buttons Or DefaultButton Or MessageIcon
    ' select button styles
    Select Case m_Params.dwStyle And &H7&
        Case [OK Only]
            blnHook = LenB(m_ButtonOK) <> 0&
        Case [OK / Cancel]
            blnHook = (LenB(m_ButtonOK) Or LenB(m_ButtonCancel)) <> 0&
        Case [Abort / Retry / Ignore]
            blnHook = (LenB(m_ButtonAbort) Or LenB(m_ButtonRetry) Or LenB(m_ButtonIgnore)) <> 0&
        Case [Yes / No / Cancel]
            blnHook = (LenB(m_ButtonYes) Or LenB(m_ButtonNo) Or LenB(m_ButtonCancel)) <> 0&
        Case [Yes / No]
            blnHook = (LenB(m_ButtonYes) Or LenB(m_ButtonNo)) <> 0&
        Case [Retry / Cancel]
            blnHook = (LenB(m_ButtonRetry) Or LenB(m_ButtonCancel)) <> 0&
        Case [Cancel / Try / Continue]
            blnHook = (LenB(m_ButtonCancel) Or LenB(m_ButtonTry) Or LenB(m_ButtonContinue)) <> 0&
    End Select
    ' temporarily replace icon?
    If WindowIcon Then lngIcon = m_Icon: m_Icon = WindowIcon
    ' hook only if we must
    If blnHook Then SetHook
    ' show the message box
    Show = MessageBoxIndirectW(m_Params)
    ' remove the hook
    RemoveHook
    ' restore original icon?
    If WindowIcon Then m_Icon = lngIcon
    ' remove icon style
    m_Params.dwStyle = m_Params.dwStyle And Not Buttons And Not DefaultButton And Not MessageIcon
    ' raise an event
    Select Case Show
        Case [Result Abort]
            RaiseEvent Result([Result Abort], m_ButtonAbort)
        Case [Result Cancel]
            RaiseEvent Result([Result Cancel], m_ButtonCancel)
        Case [Result Continue]
            RaiseEvent Result([Result Continue], m_ButtonContinue)
        Case [Result Ignore]
            RaiseEvent Result([Result Ignore], m_ButtonIgnore)
        Case [Result No]
            RaiseEvent Result([Result No], m_ButtonNo)
        Case [Result OK]
            RaiseEvent Result([Result OK], m_ButtonOK)
        Case [Result Retry]
            RaiseEvent Result([Result Retry], m_ButtonRetry)
        Case [Result Try Again]
            RaiseEvent Result([Result Try Again], m_ButtonTry)
        Case [Result Yes]
            RaiseEvent Result([Result Yes], m_ButtonYes)
    End Select
End Function
' ordinal #1, must be the last procedure!
Private Function WndMsgBoxHookProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = HCBT_ACTIVATE Then
        ' set icon?
        If m_Icon Then SendMessageW wParam, WM_SETICON, 0&, ByVal m_Icon
        ' select button styles
        Select Case m_Params.dwStyle And &H7&
            Case [OK Only]
                If LenB(m_ButtonOK) Then SetDlgItemTextW wParam, vbOK, StrPtr(m_ButtonOK)
            Case [OK / Cancel]
                If LenB(m_ButtonOK) Then SetDlgItemTextW wParam, vbOK, StrPtr(m_ButtonOK)
                If LenB(m_ButtonCancel) Then SetDlgItemTextW wParam, vbCancel, StrPtr(m_ButtonCancel)
            Case [Abort / Retry / Ignore]
                If LenB(m_ButtonAbort) Then SetDlgItemTextW wParam, vbAbort, StrPtr(m_ButtonAbort)
                If LenB(m_ButtonRetry) Then SetDlgItemTextW wParam, vbRetry, StrPtr(m_ButtonRetry)
                If LenB(m_ButtonIgnore) Then SetDlgItemTextW wParam, vbIgnore, StrPtr(m_ButtonIgnore)
            Case [Yes / No / Cancel]
                If LenB(m_ButtonYes) Then SetDlgItemTextW wParam, vbYes, StrPtr(m_ButtonYes)
                If LenB(m_ButtonNo) Then SetDlgItemTextW wParam, vbNo, StrPtr(m_ButtonNo)
                If LenB(m_ButtonCancel) Then SetDlgItemTextW wParam, vbCancel, StrPtr(m_ButtonCancel)
            Case [Yes / No]
                If LenB(m_ButtonYes) Then SetDlgItemTextW wParam, vbYes, StrPtr(m_ButtonYes)
                If LenB(m_ButtonNo) Then SetDlgItemTextW wParam, vbNo, StrPtr(m_ButtonNo)
            Case [Retry / Cancel]
                If LenB(m_ButtonRetry) Then SetDlgItemTextW wParam, vbRetry, StrPtr(m_ButtonRetry)
                If LenB(m_ButtonCancel) Then SetDlgItemTextW wParam, vbCancel, StrPtr(m_ButtonCancel)
            Case [Cancel / Try / Continue]
                If LenB(m_ButtonCancel) Then SetDlgItemTextW wParam, vbCancel, StrPtr(m_ButtonCancel)
                If LenB(m_ButtonTry) Then SetDlgItemTextW wParam, &HA, StrPtr(m_ButtonTry)
                If LenB(m_ButtonContinue) Then SetDlgItemTextW wParam, &HB, StrPtr(m_ButtonContinue)
        End Select
        ' we must not call RemoveHook here; instead just end hook
        If m_Hook Then UnhookWindowsHookEx m_Hook: m_Hook = 0
    End If
End Function

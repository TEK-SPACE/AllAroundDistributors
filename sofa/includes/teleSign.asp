<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: Telesign Functions
%>

<%
Function callRequest(uri, custID, countryCode, phoneNum, language, randCode, priority, delayTime, redialCount)
    Dim strPost, httpPost
    
    pXmlHTTPComponent		= getSettingKey("pXmlHTTPComponent")
    
    strPost = "customerID=" & custID & _
            "&country_code="& countryCode & _ 
            "&phone_num="& phoneNum & _ 
            "&language="& language & _ 
            "&verify_code="& randCode & _ 
            "&priority="& priority & _ 
            "&delay_time="& delayTime & _ 
            "&redial_count="& redialCount     

if pXmlHTTPComponent="Microsoft.XMLHTTP" then
 set httppost = server.Createobject("Microsoft.XMLHTTP")
end if

if pXmlHTTPComponent="Msxml2.ServerXMLHTTP.5.0" then
 set httppost = server.Createobject("Msxml2.ServerXMLHTTP.5.0")
end if

if pXmlHTTPComponent="Msxml2.ServerXMLHTTP.4.0" then
 set httppost = server.Createobject("Msxml2.ServerXMLHTTP.4.0")
end if

if pXmlHTTPComponent="Msxml2.ServerXMLHTTP" then
 set httppost = server.Createobject("Msxml2.ServerXMLHTTP")
end if

if pXmlHTTPComponent="NONE" then
 response.redirect "comersus_message.asp?message="&Server.UrlEncode("Please configure xmlhttp component or disable this feature.")	 
end if
    
    
    httpPost.Open "POST", uri, false
    httpPost.SetRequestHeader "Content-type","application/x-www-form-urlencoded"
    httpPost.SetRequestHeader "Content-length", Len(strPost)
    httpPost.Send(strPost)
    callRequest = httpPost.responseText
    set httpPost = Nothing
End function

Function callStatus(uri, custID, refID, verifyCode)

pXmlHTTPComponent		= getSettingKey("pXmlHTTPComponent")

    Dim strPost, httpPost
    strPost = "customerID=" & custID & _
            "&refID="& refID & _ 
            "&verify_code="& verifyCode 
            

if pXmlHTTPComponent="Microsoft.XMLHTTP" then
 set httppost = server.Createobject("Microsoft.XMLHTTP")
end if

if pXmlHTTPComponent="Msxml2.ServerXMLHTTP.5.0" then
 set httppost = server.Createobject("Msxml2.ServerXMLHTTP.5.0")
end if

if pXmlHTTPComponent="Msxml2.ServerXMLHTTP.4.0" then
 set httppost = server.Createobject("Msxml2.ServerXMLHTTP.4.0")
end if

if pXmlHTTPComponent="Msxml2.ServerXMLHTTP" then
 set httppost = server.Createobject("Msxml2.ServerXMLHTTP")
end if
    
    
    httpPost.Open "POST", uri, false
    httpPost.SetRequestHeader "Content-type","application/x-www-form-urlencoded"
    httpPost.SetRequestHeader "Content-length", Len(strPost)
    httpPost.Send(strPost)
    callStatus = httpPost.responseText
    set httpPost = Nothing
End function

Function GetCountryPhoneCode(pCountryCode)
    Dim arrCountryPhoneCode(170,3)
    'column 1: PhoneCode
    'column 2: db CountryCode
    'column 3: Country Description
    
    arrCountryPhoneCode(1,3)  ="UNITED STATES                                 "    
    arrCountryPhoneCode(2,3)  ="KAZAKHSTAN                                    "    
    arrCountryPhoneCode(3,3)  ="RUSSIAN FEDERATION                            "    
    arrCountryPhoneCode(4,3)  ="EGYPT                                         "    
    arrCountryPhoneCode(5,3)  ="SOUTH AFRICA                                  "    
    arrCountryPhoneCode(6,3)  ="GREECE                                        "    
    arrCountryPhoneCode(7,3)  ="NETHERLANDS                                   "    
    arrCountryPhoneCode(8,3)  ="BELGIUM                                       "    
    arrCountryPhoneCode(9,3)  ="FRANCE                                        "    
    arrCountryPhoneCode(10,3) ="SPAIN                                        "    
    arrCountryPhoneCode(11,3) ="HUNGARY                                      "    
    arrCountryPhoneCode(12,3) ="ITALY                                        "    
    arrCountryPhoneCode(13,3) ="ROMANIA                                      "    
    arrCountryPhoneCode(14,3) ="SWITZERLAND                                  "    
    arrCountryPhoneCode(15,3) ="AUSTRIA                                      "    
    arrCountryPhoneCode(16,3) ="UNITED KINGDOM                               "
    arrCountryPhoneCode(17,3) ="DENMARK                                      "
    arrCountryPhoneCode(18,3) ="SWEDEN                                       "
    arrCountryPhoneCode(19,3) ="NORWAY                                       "
    arrCountryPhoneCode(20,3) ="POLAND                                       "
    arrCountryPhoneCode(21,3) ="GERMANY                                      "
    arrCountryPhoneCode(22,3) ="PERU                                         "
    arrCountryPhoneCode(23,3) ="MEXICO                                       "
    arrCountryPhoneCode(24,3) ="CUBA                                         "
    arrCountryPhoneCode(25,3) ="ARGENTINA                                    "
    arrCountryPhoneCode(26,3) ="BRAZIL                                       "
    arrCountryPhoneCode(27,3) ="CHILE                                        "
    arrCountryPhoneCode(28,3) ="COLOMBIA                                     "
    arrCountryPhoneCode(29,3) ="VENEZUELA                                    "
    arrCountryPhoneCode(30,3) ="MALAYSIA                                     "
    arrCountryPhoneCode(31,3) ="AUSTRALIA                                    "
    arrCountryPhoneCode(32,3) ="INDONESIA                                    "
    arrCountryPhoneCode(33,3) ="PHILIPPINES                                  "
    arrCountryPhoneCode(34,3) ="NEW ZEALAND                                  "
    arrCountryPhoneCode(35,3) ="SINGAPORE                                    "
    arrCountryPhoneCode(36,3) ="THAILAND                                     "
    arrCountryPhoneCode(37,3) ="JAPAN                                        "
    arrCountryPhoneCode(38,3) ="KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF       "
    arrCountryPhoneCode(39,3) ="KOREA, REPUBLIC OF                           "
    arrCountryPhoneCode(40,3) ="VIET NAM                                     "
    arrCountryPhoneCode(41,3) ="CHINA                                        "
    arrCountryPhoneCode(42,3) ="TURKEY                                       "
    arrCountryPhoneCode(43,3) ="INDIA                                        "
    arrCountryPhoneCode(44,3) ="PAKISTAN                                     "
    arrCountryPhoneCode(45,3) ="AFGHANISTAN                                  "
    arrCountryPhoneCode(46,3) ="SRI LANKA                                    "
    arrCountryPhoneCode(47,3) ="IRAN (ISLAMIC REPUBLIC OF)                   "
    arrCountryPhoneCode(48,3) ="MOROCCO                                      "
    arrCountryPhoneCode(49,3) ="ALGERIA                                      "
    arrCountryPhoneCode(50,3) ="TUNISIA                                      "
    arrCountryPhoneCode(51,3) ="LIBYAN ARAB JAMAHIRIYA                       "
    arrCountryPhoneCode(52,3) ="GAMBIA                                       "
    arrCountryPhoneCode(53,3) ="SENEGAL                                      "
    arrCountryPhoneCode(54,3) ="MALI                                         "
    arrCountryPhoneCode(55,3) ="BURKINA FASO                                 "
    arrCountryPhoneCode(56,3) ="NIGER                                        "
    arrCountryPhoneCode(57,3) ="TOGO                                         "
    arrCountryPhoneCode(58,3) ="BENIN                                        "
    arrCountryPhoneCode(59,3) ="MAURITIUS                                    "
    arrCountryPhoneCode(60,3) ="SIERRA LEONE                                 "
    arrCountryPhoneCode(61,3) ="GHANA                                        "
    arrCountryPhoneCode(62,3) ="NIGERIA                                      "
    arrCountryPhoneCode(63,3) ="CHAD                                         "
    arrCountryPhoneCode(64,3) ="CENTRAL AFRICAN REPUBLIC                     "
    arrCountryPhoneCode(65,3) ="CAMEROON                                     "
    arrCountryPhoneCode(66,3) ="CAPE VERDE                                   "
    arrCountryPhoneCode(67,3) ="EQUATORIAL GUINEA                            "
    arrCountryPhoneCode(68,3) ="GABON                                        "
    arrCountryPhoneCode(69,3) ="CONGO                                        "
    arrCountryPhoneCode(70,3) ="ANGOLA                                       "
    arrCountryPhoneCode(71,3) ="SEYCHELLES                                   "
    arrCountryPhoneCode(72,3) ="SUDAN                                        "
    arrCountryPhoneCode(73,3) ="RWANDA                                       "
    arrCountryPhoneCode(74,3) ="ETHIOPIA                                     "
    arrCountryPhoneCode(75,3) ="KENYA                                        "
    arrCountryPhoneCode(76,3) ="TANZANIA, UNITED REPUBLIC OF                 "
    arrCountryPhoneCode(77,3) ="UGANDA                                       "
    arrCountryPhoneCode(78,3) ="BURUNDI                                      "
    arrCountryPhoneCode(79,3) ="MOZAMBIQUE                                   "
    arrCountryPhoneCode(80,3) ="ZAMBIA                                       "
    arrCountryPhoneCode(81,3) ="MADAGASCAR                                   "
    arrCountryPhoneCode(82,3) ="REUNION                                      "
    arrCountryPhoneCode(83,3) ="NAMIBIA                                      "
    arrCountryPhoneCode(84,3) ="MALAWI                                       "
    arrCountryPhoneCode(85,3) ="LESOTHO                                      "
    arrCountryPhoneCode(86,3) ="BOTSWANA                                     "
    arrCountryPhoneCode(87,3) ="SWAZILAND                                    "
    arrCountryPhoneCode(88,3) ="ARUBA                                        "
    arrCountryPhoneCode(89,3) ="FALKLAND ISLANDS (MALVINAS)                  "
    arrCountryPhoneCode(90,3) ="GREENLAND                                    "
    arrCountryPhoneCode(91,3) ="GIBRALTAR                                    "
    arrCountryPhoneCode(92,3) ="PORTUGAL                                     "
    arrCountryPhoneCode(93,3) ="LUXEMBOURG                                   "
    arrCountryPhoneCode(94,3) ="IRELAND                                      "
    arrCountryPhoneCode(95,3) ="ICELAND                                      "
    arrCountryPhoneCode(96,3) ="ALBANIA                                      "
    arrCountryPhoneCode(97,3) ="MALTA                                        "
    arrCountryPhoneCode(98,3) ="CYPRUS                                       "
    arrCountryPhoneCode(99,3) ="FINLAND                                      "
    arrCountryPhoneCode(100,3)="BULGARIA                                    "
    arrCountryPhoneCode(101,3)="LITHUANIA                                   "
    arrCountryPhoneCode(102,3)="LATVIA                                      "
    arrCountryPhoneCode(103,3)="ESTONIA                                     "
    arrCountryPhoneCode(104,3)="MOLDOVA, REPUBLIC OF                        "
    arrCountryPhoneCode(105,3)="ARMENIA                                     "
    arrCountryPhoneCode(106,3)="BELARUS                                     "
    arrCountryPhoneCode(107,3)="ANDORRA                                     "
    arrCountryPhoneCode(108,3)="MONACO                                      "
    arrCountryPhoneCode(109,3)="UKRAINE                                     "
    arrCountryPhoneCode(110,3)="YUGOSLAVIA                                  "
    arrCountryPhoneCode(111,3)="CROATIA                                     "
    arrCountryPhoneCode(112,3)="SLOVENIA                                    "
    arrCountryPhoneCode(113,3)="BOSNIA HERCEGOVINA                          "
    arrCountryPhoneCode(114,3)="CZECH REPUBLIC                              "
    arrCountryPhoneCode(115,3)="SLOVAKIA                                    "
    arrCountryPhoneCode(116,3)="LIECHTENSTEIN                               "
    arrCountryPhoneCode(117,3)="BELIZE                                      "
    arrCountryPhoneCode(118,3)="EL SALVADOR                                 "
    arrCountryPhoneCode(119,3)="PANAMA                                      "
    arrCountryPhoneCode(120,3)="GUADELOUPE                                  "
    arrCountryPhoneCode(121,3)="BOLIVIA                                     "
    arrCountryPhoneCode(122,3)="GUYANA                                      "
    arrCountryPhoneCode(123,3)="ECUADOR                                     "
    arrCountryPhoneCode(124,3)="FRENCH GUIANA                               "
    arrCountryPhoneCode(125,3)="PARAGUAY                                    "
    arrCountryPhoneCode(126,3)="MARTINIQUE                                  "
    arrCountryPhoneCode(127,3)="SURINAME                                    "
    arrCountryPhoneCode(128,3)="URUGUAY                                     "
    arrCountryPhoneCode(129,3)="NETHERLANDS ANTILLES                        "
    arrCountryPhoneCode(130,3)="BRUNEI DARUSSALAM                           "
    arrCountryPhoneCode(131,3)="VANUATU                                     "
    arrCountryPhoneCode(132,3)="FIJI                                        "
    arrCountryPhoneCode(133,3)="NEW CALEDONIA                               "
    arrCountryPhoneCode(134,3)="FRENCH POLYNESIA                            "
    arrCountryPhoneCode(135,3)="HONG KONG                                   "
    arrCountryPhoneCode(136,3)="MACAU                                       "
    arrCountryPhoneCode(137,3)="CAMBODIA                                    "
    arrCountryPhoneCode(138,3)="LAO PEOPLE'S DEMOCRATIC REPUBLIC            "
    arrCountryPhoneCode(139,3)="BANGLADESH                                  "
    arrCountryPhoneCode(140,3)="TAIWAN, REPUBLIC OF CHINA                   "
    arrCountryPhoneCode(141,3)="MALDIVES                                    "
    arrCountryPhoneCode(142,3)="LEBANON                                     "
    arrCountryPhoneCode(143,3)="JORDAN                                      "
    arrCountryPhoneCode(144,3)="SYRIAN ARAB REPUBLIC                        "
    arrCountryPhoneCode(145,3)="IRAQ                                        "
    arrCountryPhoneCode(146,3)="KUWAIT                                      "
    arrCountryPhoneCode(147,3)="SAUDI ARABIA                                "
    arrCountryPhoneCode(148,3)="YEMEN, REPUBLIC OF                          "
    arrCountryPhoneCode(149,3)="OMAN                                        "
    arrCountryPhoneCode(150,3)="UNITED ARAB EMIRATES                        "
    arrCountryPhoneCode(151,3)="ISRAEL                                      "
    arrCountryPhoneCode(152,3)="BAHRAIN                                     "
    arrCountryPhoneCode(153,3)="QATAR                                       "
    arrCountryPhoneCode(154,3)="MONGOLIA                                    "
    arrCountryPhoneCode(155,3)="NEPAL                                       "
    arrCountryPhoneCode(156,3)="TAJIKISTAN                                  "
    arrCountryPhoneCode(157,3)="TURKMENISTAN                                "
    arrCountryPhoneCode(158,3)="AZERBAIJAN                                  "
    arrCountryPhoneCode(159,3)="GEORGIA                                     "
    arrCountryPhoneCode(160,3)="KYRGYZSTAN                                  "
    arrCountryPhoneCode(161,3)="UZBEKISTAN                                  "
    arrCountryPhoneCode(162,3)="ANTIGUA AND BARBUDA                         "
    arrCountryPhoneCode(163,3)="CAYMAN ISLANDS                              "
    arrCountryPhoneCode(164,3)="BERMUDA                                     "
    arrCountryPhoneCode(165,3)="PUERTO RICO                                 "
    arrCountryPhoneCode(1,2)  ="US"   
    arrCountryPhoneCode(2,2)  ="KZ"   
    arrCountryPhoneCode(3,2)  ="RU"   
    arrCountryPhoneCode(4,2)  ="EG"   
    arrCountryPhoneCode(5,2)  ="ZA"   
    arrCountryPhoneCode(6,2)  ="GR"   
    arrCountryPhoneCode(7,2)  ="NL"   
    arrCountryPhoneCode(8,2)  ="BE"   
    arrCountryPhoneCode(9,2)  ="FR"   
    arrCountryPhoneCode(10,2) ="ES"   
    arrCountryPhoneCode(11,2) ="HU"   
    arrCountryPhoneCode(12,2) ="IT"   
    arrCountryPhoneCode(13,2) ="RO"   
    arrCountryPhoneCode(14,2) ="CH"   
    arrCountryPhoneCode(15,2) ="AT"   
    arrCountryPhoneCode(16,2) ="GB"   
    arrCountryPhoneCode(17,2) ="DK"   
    arrCountryPhoneCode(18,2) ="SE"   
    arrCountryPhoneCode(19,2) ="NO"   
    arrCountryPhoneCode(20,2) ="PL"   
    arrCountryPhoneCode(21,2) ="DE"   
    arrCountryPhoneCode(22,2) ="PE"   
    arrCountryPhoneCode(23,2) ="MX"   
    arrCountryPhoneCode(24,2) ="CU"   
    arrCountryPhoneCode(25,2) ="AR"   
    arrCountryPhoneCode(26,2) ="BR"   
    arrCountryPhoneCode(27,2) ="CL"   
    arrCountryPhoneCode(28,2) ="CO"   
    arrCountryPhoneCode(29,2) ="VE"   
    arrCountryPhoneCode(30,2) ="MY"   
    arrCountryPhoneCode(31,2) ="AU"   
    arrCountryPhoneCode(32,2) ="ID"   
    arrCountryPhoneCode(33,2) ="PH"   
    arrCountryPhoneCode(34,2) ="NZ"   
    arrCountryPhoneCode(35,2) ="SG"   
    arrCountryPhoneCode(36,2) ="TH"   
    arrCountryPhoneCode(37,2) ="JP"   
    arrCountryPhoneCode(38,2) ="KP"   
    arrCountryPhoneCode(39,2) ="KR"   
    arrCountryPhoneCode(40,2) ="VN"   
    arrCountryPhoneCode(41,2) ="CN"   
    arrCountryPhoneCode(42,2) ="TR"   
    arrCountryPhoneCode(43,2) ="IN"   
    arrCountryPhoneCode(44,2) ="PK"   
    arrCountryPhoneCode(45,2) ="AF"   
    arrCountryPhoneCode(46,2) ="LK"   
    arrCountryPhoneCode(47,2) ="IR"   
    arrCountryPhoneCode(48,2) ="MA"   
    arrCountryPhoneCode(49,2) ="DZ"   
    arrCountryPhoneCode(50,2) ="TN"   
    arrCountryPhoneCode(51,2) ="LY"   
    arrCountryPhoneCode(52,2) ="GM"   
    arrCountryPhoneCode(53,2) ="SN"   
    arrCountryPhoneCode(54,2) ="ML"   
    arrCountryPhoneCode(55,2) ="BF"   
    arrCountryPhoneCode(56,2) ="NE"   
    arrCountryPhoneCode(57,2) ="TG"   
    arrCountryPhoneCode(58,2) ="BJ"   
    arrCountryPhoneCode(59,2) ="MU"   
    arrCountryPhoneCode(60,2) ="SL"   
    arrCountryPhoneCode(61,2) ="GH"   
    arrCountryPhoneCode(62,2) ="NG"   
    arrCountryPhoneCode(63,2) ="TD"   
    arrCountryPhoneCode(64,2) ="CF"   
    arrCountryPhoneCode(65,2) ="CM"   
    arrCountryPhoneCode(66,2) ="CV"   
    arrCountryPhoneCode(67,2) ="GQ"   
    arrCountryPhoneCode(68,2) ="GA"   
    arrCountryPhoneCode(69,2) ="CG"   
    arrCountryPhoneCode(70,2) ="AO"   
    arrCountryPhoneCode(71,2) ="SC"   
    arrCountryPhoneCode(72,2) ="SD"   
    arrCountryPhoneCode(73,2) ="RW"   
    arrCountryPhoneCode(74,2) ="ET"   
    arrCountryPhoneCode(75,2) ="KE"   
    arrCountryPhoneCode(76,2) ="TZ"   
    arrCountryPhoneCode(77,2) ="UG"   
    arrCountryPhoneCode(78,2) ="BI"   
    arrCountryPhoneCode(79,2) ="MZ"   
    arrCountryPhoneCode(80,2) ="ZM"   
    arrCountryPhoneCode(81,2) ="MG"   
    arrCountryPhoneCode(82,2) ="RE"   
    arrCountryPhoneCode(83,2) ="NA"   
    arrCountryPhoneCode(84,2) ="MW"   
    arrCountryPhoneCode(85,2) ="LS"   
    arrCountryPhoneCode(86,2) ="BW"   
    arrCountryPhoneCode(87,2) ="SZ"   
    arrCountryPhoneCode(88,2) ="AW"   
    arrCountryPhoneCode(89,2) ="FK"   
    arrCountryPhoneCode(90,2) ="GL"   
    arrCountryPhoneCode(91,2) ="GI"   
    arrCountryPhoneCode(92,2) ="PT"   
    arrCountryPhoneCode(93,2) ="LU"   
    arrCountryPhoneCode(94,2) ="IE"   
    arrCountryPhoneCode(95,2) ="IS"   
    arrCountryPhoneCode(96,2) ="AL"   
    arrCountryPhoneCode(97,2) ="MT"   
    arrCountryPhoneCode(98,2) ="CY"   
    arrCountryPhoneCode(99,2) ="FI"   
    arrCountryPhoneCode(100,2)="BG"   
    arrCountryPhoneCode(101,2)="LT"   
    arrCountryPhoneCode(102,2)="LV"   
    arrCountryPhoneCode(103,2)="EE"   
    arrCountryPhoneCode(104,2)="MD"   
    arrCountryPhoneCode(105,2)="AM"   
    arrCountryPhoneCode(106,2)="BY"   
    arrCountryPhoneCode(107,2)="AD"   
    arrCountryPhoneCode(108,2)="MC"   
    arrCountryPhoneCode(109,2)="UA"   
    arrCountryPhoneCode(110,2)="YU"   
    arrCountryPhoneCode(111,2)="HR"   
    arrCountryPhoneCode(112,2)="SI"   
    arrCountryPhoneCode(113,2)="BA"   
    arrCountryPhoneCode(114,2)="CZ"   
    arrCountryPhoneCode(115,2)="SK"   
    arrCountryPhoneCode(116,2)="LI"   
    arrCountryPhoneCode(117,2)="BZ"   
    arrCountryPhoneCode(118,2)="SV"   
    arrCountryPhoneCode(119,2)="PA"   
    arrCountryPhoneCode(120,2)="GP"   
    arrCountryPhoneCode(121,2)="BO"   
    arrCountryPhoneCode(122,2)="GY"   
    arrCountryPhoneCode(123,2)="EC"   
    arrCountryPhoneCode(124,2)="GF"   
    arrCountryPhoneCode(125,2)="PY"   
    arrCountryPhoneCode(126,2)="MQ"   
    arrCountryPhoneCode(127,2)="SR"   
    arrCountryPhoneCode(128,2)="UY"   
    arrCountryPhoneCode(129,2)="AN"   
    arrCountryPhoneCode(130,2)="BN"   
    arrCountryPhoneCode(131,2)="VU"   
    arrCountryPhoneCode(132,2)="FJ"   
    arrCountryPhoneCode(133,2)="NC"   
    arrCountryPhoneCode(134,2)="PF"   
    arrCountryPhoneCode(135,2)="HK"   
    arrCountryPhoneCode(136,2)="MO"   
    arrCountryPhoneCode(137,2)="KH"   
    arrCountryPhoneCode(138,2)="LA"   
    arrCountryPhoneCode(139,2)="BD"   
    arrCountryPhoneCode(140,2)="TW"   
    arrCountryPhoneCode(141,2)="MV"   
    arrCountryPhoneCode(142,2)="LB"   
    arrCountryPhoneCode(143,2)="JO"   
    arrCountryPhoneCode(144,2)="SY"   
    arrCountryPhoneCode(145,2)="IQ"   
    arrCountryPhoneCode(146,2)="KW"   
    arrCountryPhoneCode(147,2)="SA"   
    arrCountryPhoneCode(148,2)="YE"   
    arrCountryPhoneCode(149,2)="OM"   
    arrCountryPhoneCode(150,2)="AE"   
    arrCountryPhoneCode(151,2)="IL"   
    arrCountryPhoneCode(152,2)="BH"   
    arrCountryPhoneCode(153,2)="QA"   
    arrCountryPhoneCode(154,2)="MN"   
    arrCountryPhoneCode(155,2)="NP"   
    arrCountryPhoneCode(156,2)="TJ"   
    arrCountryPhoneCode(157,2)="TM"   
    arrCountryPhoneCode(158,2)="AZ"   
    arrCountryPhoneCode(159,2)="GE"   
    arrCountryPhoneCode(160,2)="KG"   
    arrCountryPhoneCode(161,2)="UZ"   
    arrCountryPhoneCode(162,2)="AG"   
    arrCountryPhoneCode(163,2)="KY"   
    arrCountryPhoneCode(164,2)="BM"   
    arrCountryPhoneCode(165,2)="PR"   
    arrCountryPhoneCode(1,1)  = 1   
    arrCountryPhoneCode(2,1)  = 7   
    arrCountryPhoneCode(3,1)  = 7   
    arrCountryPhoneCode(4,1)  = 20  
    arrCountryPhoneCode(5,1)  = 27  
    arrCountryPhoneCode(6,1)  = 30  
    arrCountryPhoneCode(7,1)  = 31  
    arrCountryPhoneCode(8,1)  = 32  
    arrCountryPhoneCode(9,1)  = 33  
    arrCountryPhoneCode(10,1) = 34  
    arrCountryPhoneCode(11,1) = 36  
    arrCountryPhoneCode(12,1) = 39  
    arrCountryPhoneCode(13,1) = 40  
    arrCountryPhoneCode(14,1) = 41  
    arrCountryPhoneCode(15,1) = 43  
    arrCountryPhoneCode(16,1) = 44  
    arrCountryPhoneCode(17,1) = 45  
    arrCountryPhoneCode(18,1) = 46  
    arrCountryPhoneCode(19,1) = 47  
    arrCountryPhoneCode(20,1) = 48  
    arrCountryPhoneCode(21,1) = 49  
    arrCountryPhoneCode(22,1) = 51  
    arrCountryPhoneCode(23,1) = 52  
    arrCountryPhoneCode(24,1) = 53  
    arrCountryPhoneCode(25,1) = 54  
    arrCountryPhoneCode(26,1) = 55  
    arrCountryPhoneCode(27,1) = 56  
    arrCountryPhoneCode(28,1) = 57  
    arrCountryPhoneCode(29,1) = 58  
    arrCountryPhoneCode(30,1) = 60  
    arrCountryPhoneCode(31,1) = 61  
    arrCountryPhoneCode(32,1) = 62  
    arrCountryPhoneCode(33,1) = 63  
    arrCountryPhoneCode(34,1) = 64  
    arrCountryPhoneCode(35,1) = 65  
    arrCountryPhoneCode(36,1) = 66  
    arrCountryPhoneCode(37,1) = 81  
    arrCountryPhoneCode(38,1) = 82  
    arrCountryPhoneCode(39,1) = 82  
    arrCountryPhoneCode(40,1) = 84  
    arrCountryPhoneCode(41,1) = 86  
    arrCountryPhoneCode(42,1) = 90  
    arrCountryPhoneCode(43,1) = 91  
    arrCountryPhoneCode(44,1) = 92  
    arrCountryPhoneCode(45,1) = 93  
    arrCountryPhoneCode(46,1) = 94  
    arrCountryPhoneCode(47,1) = 98  
    arrCountryPhoneCode(48,1) = 212 
    arrCountryPhoneCode(49,1) = 213 
    arrCountryPhoneCode(50,1) = 216 
    arrCountryPhoneCode(51,1) = 218 
    arrCountryPhoneCode(52,1) = 220 
    arrCountryPhoneCode(53,1) = 221 
    arrCountryPhoneCode(54,1) = 223 
    arrCountryPhoneCode(55,1) = 226 
    arrCountryPhoneCode(56,1) = 227 
    arrCountryPhoneCode(57,1) = 228 
    arrCountryPhoneCode(58,1) = 229 
    arrCountryPhoneCode(59,1) = 230 
    arrCountryPhoneCode(60,1) = 232 
    arrCountryPhoneCode(61,1) = 233 
    arrCountryPhoneCode(62,1) = 234 
    arrCountryPhoneCode(63,1) = 235 
    arrCountryPhoneCode(64,1) = 236 
    arrCountryPhoneCode(65,1) = 237 
    arrCountryPhoneCode(66,1) = 238 
    arrCountryPhoneCode(67,1) = 240 
    arrCountryPhoneCode(68,1) = 241 
    arrCountryPhoneCode(69,1) = 243 
    arrCountryPhoneCode(70,1) = 244 
    arrCountryPhoneCode(71,1) = 248 
    arrCountryPhoneCode(72,1) = 249 
    arrCountryPhoneCode(73,1) = 250 
    arrCountryPhoneCode(74,1) = 251 
    arrCountryPhoneCode(75,1) = 254 
    arrCountryPhoneCode(76,1) = 255 
    arrCountryPhoneCode(77,1) = 256 
    arrCountryPhoneCode(78,1) = 257 
    arrCountryPhoneCode(79,1) = 258 
    arrCountryPhoneCode(80,1) = 260 
    arrCountryPhoneCode(81,1) = 261 
    arrCountryPhoneCode(82,1) = 262 
    arrCountryPhoneCode(83,1) = 264 
    arrCountryPhoneCode(84,1) = 265 
    arrCountryPhoneCode(85,1) = 266 
    arrCountryPhoneCode(86,1) = 267 
    arrCountryPhoneCode(87,1) = 268 
    arrCountryPhoneCode(88,1) = 297 
    arrCountryPhoneCode(89,1) = 298 
    arrCountryPhoneCode(90,1) = 299 
    arrCountryPhoneCode(91,1) = 350 
    arrCountryPhoneCode(92,1) = 351 
    arrCountryPhoneCode(93,1) = 352 
    arrCountryPhoneCode(94,1) = 353 
    arrCountryPhoneCode(95,1) = 354 
    arrCountryPhoneCode(96,1) = 355 
    arrCountryPhoneCode(97,1) = 356 
    arrCountryPhoneCode(98,1) = 357 
    arrCountryPhoneCode(99,1) = 358 
    arrCountryPhoneCode(100,1)= 359 
    arrCountryPhoneCode(101,1)= 370 
    arrCountryPhoneCode(102,1)= 371 
    arrCountryPhoneCode(103,1)= 372 
    arrCountryPhoneCode(104,1)= 373 
    arrCountryPhoneCode(105,1)= 374 
    arrCountryPhoneCode(106,1)= 375 
    arrCountryPhoneCode(107,1)= 376 
    arrCountryPhoneCode(108,1)= 377 
    arrCountryPhoneCode(109,1)= 380 
    arrCountryPhoneCode(110,1)= 381 
    arrCountryPhoneCode(111,1)= 385 
    arrCountryPhoneCode(112,1)= 386 
    arrCountryPhoneCode(113,1)= 387 
    arrCountryPhoneCode(114,1)= 420 
    arrCountryPhoneCode(115,1)= 421 
    arrCountryPhoneCode(116,1)= 423 
    arrCountryPhoneCode(117,1)= 501 
    arrCountryPhoneCode(118,1)= 503 
    arrCountryPhoneCode(119,1)= 507 
    arrCountryPhoneCode(120,1)= 590 
    arrCountryPhoneCode(121,1)= 591 
    arrCountryPhoneCode(122,1)= 592 
    arrCountryPhoneCode(123,1)= 593 
    arrCountryPhoneCode(124,1)= 594 
    arrCountryPhoneCode(125,1)= 595 
    arrCountryPhoneCode(126,1)= 596 
    arrCountryPhoneCode(127,1)= 597 
    arrCountryPhoneCode(128,1)= 598 
    arrCountryPhoneCode(129,1)= 599 
    arrCountryPhoneCode(130,1)= 673 
    arrCountryPhoneCode(131,1)= 678 
    arrCountryPhoneCode(132,1)= 679 
    arrCountryPhoneCode(133,1)= 687 
    arrCountryPhoneCode(134,1)= 689 
    arrCountryPhoneCode(135,1)= 852 
    arrCountryPhoneCode(136,1)= 853 
    arrCountryPhoneCode(137,1)= 855 
    arrCountryPhoneCode(138,1)= 856 
    arrCountryPhoneCode(139,1)= 880 
    arrCountryPhoneCode(140,1)= 886 
    arrCountryPhoneCode(141,1)= 960 
    arrCountryPhoneCode(142,1)= 961 
    arrCountryPhoneCode(143,1)= 962 
    arrCountryPhoneCode(144,1)= 963 
    arrCountryPhoneCode(145,1)= 964 
    arrCountryPhoneCode(146,1)= 965 
    arrCountryPhoneCode(147,1)= 966 
    arrCountryPhoneCode(148,1)= 967 
    arrCountryPhoneCode(149,1)= 968 
    arrCountryPhoneCode(150,1)= 971 
    arrCountryPhoneCode(151,1)= 972 
    arrCountryPhoneCode(152,1)= 973 
    arrCountryPhoneCode(153,1)= 974 
    arrCountryPhoneCode(154,1)= 976 
    arrCountryPhoneCode(155,1)= 977 
    arrCountryPhoneCode(156,1)= 992 
    arrCountryPhoneCode(157,1)= 993 
    arrCountryPhoneCode(158,1)= 994 
    arrCountryPhoneCode(159,1)= 995 
    arrCountryPhoneCode(160,1)= 996 
    arrCountryPhoneCode(161,1)= 998 
    arrCountryPhoneCode(162,1)= 1268
    arrCountryPhoneCode(163,1)= 1345
    arrCountryPhoneCode(164,1)= 1441
    arrCountryPhoneCode(165,1)= 1787
    pCode = 0
    For x = 1 to 170
        if trim(arrCountryPhoneCode(x,2)) = trim(pCountryCode) then
            pCode = arrCountryPhoneCode(x,1)
        end if
    Next    
    GetCountryPhoneCode = pCode
                
End function
%>

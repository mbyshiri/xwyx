Function GetSpellingChar(oValue,isAll)
Dim objDic
Set objDic = CreateObject("Scripting.Dictionary")
objDic.Add "a", -20319
objDic.Add "ai", -20317
objDic.Add "an", -20304
objDic.Add "ang", -20295
objDic.Add "ao", -20292
objDic.Add "ba", -20283
objDic.Add "bai", -20265
objDic.Add "ban", -20257
objDic.Add "bang", -20242
objDic.Add "bao", -20230
objDic.Add "bei", -20051
objDic.Add "ben", -20036
objDic.Add "beng", -20032
objDic.Add "bi", -20026
objDic.Add "bian", -20002
objDic.Add "biao", -19990
objDic.Add "bie", -19986
objDic.Add "bin", -19982
objDic.Add "bing", -19976
objDic.Add "bo", -19805
objDic.Add "bu", -19784
objDic.Add "ca", -19775
objDic.Add "cai", -19774
objDic.Add "can", -19763
objDic.Add "cang", -19756
objDic.Add "cao", -19751
objDic.Add "ce", -19746
objDic.Add "ceng", -19741
objDic.Add "cha", -19739
objDic.Add "chai", -19728
objDic.Add "chan", -19725
objDic.Add "chang", -19715
objDic.Add "chao", -19540
objDic.Add "che", -19531
objDic.Add "chen", -19525
objDic.Add "cheng", -19515
objDic.Add "chi", -19500
objDic.Add "chong", -19484
objDic.Add "chou", -19479
objDic.Add "chu", -19467
objDic.Add "chuai", -19289
objDic.Add "chuan", -19288
objDic.Add "chuang", -19281
objDic.Add "chui", -19275
objDic.Add "chun", -19270
objDic.Add "chuo", -19263
objDic.Add "ci", -19261
objDic.Add "cong", -19249
objDic.Add "cou", -19243
objDic.Add "cu", -19242
objDic.Add "cuan", -19238
objDic.Add "cui", -19235
objDic.Add "cun", -19227
objDic.Add "cuo", -19224
objDic.Add "da", -19218
objDic.Add "dai", -19212
objDic.Add "dan", -19038
objDic.Add "dang", -19023
objDic.Add "dao", -19018
objDic.Add "de", -19006
objDic.Add "deng", -19003
objDic.Add "di", -18996
objDic.Add "dian", -18977
objDic.Add "diao", -18961
objDic.Add "die", -18952
objDic.Add "ding", -18783
objDic.Add "diu", -18774
objDic.Add "dong", -18773
objDic.Add "dou", -18763
objDic.Add "du", -18756
objDic.Add "duan", -18741
objDic.Add "dui", -18735
objDic.Add "dun", -18731
objDic.Add "duo", -18722
objDic.Add "e", -18710
objDic.Add "en", -18697
objDic.Add "er", -18696
objDic.Add "fa", -18526
objDic.Add "fan", -18518
objDic.Add "fang", -18501
objDic.Add "fei", -18490
objDic.Add "fen", -18478
objDic.Add "feng", -18463
objDic.Add "fo", -18448
objDic.Add "fou", -18447
objDic.Add "fu", -18446
objDic.Add "ga", -18239
objDic.Add "gai", -18237
objDic.Add "gan", -18231
objDic.Add "gang", -18220
objDic.Add "gao", -18211
objDic.Add "ge", -18201
objDic.Add "gei", -18184
objDic.Add "gen", -18183
objDic.Add "geng", -18181
objDic.Add "gong", -18012
objDic.Add "gou", -17997
objDic.Add "gu", -17988
objDic.Add "gua", -17970
objDic.Add "guai", -17964
objDic.Add "guan", -17961
objDic.Add "guang", -17950
objDic.Add "gui", -17947
objDic.Add "gun", -17931
objDic.Add "guo", -17928
objDic.Add "ha", -17922
objDic.Add "hai", -17759
objDic.Add "han", -17752
objDic.Add "hang", -17733
objDic.Add "hao", -17730
objDic.Add "he", -17721
objDic.Add "hei", -17703
objDic.Add "hen", -17701
objDic.Add "heng", -17697
objDic.Add "hong", -17692
objDic.Add "hou", -17683
objDic.Add "hu", -17676
objDic.Add "hua", -17496
objDic.Add "huai", -17487
objDic.Add "huan", -17482
objDic.Add "huang", -17468
objDic.Add "hui", -17454
objDic.Add "hun", -17433
objDic.Add "huo", -17427
objDic.Add "ji", -17417
objDic.Add "jia", -17202
objDic.Add "jian", -17185
objDic.Add "jiang", -16983
objDic.Add "jiao", -16970
objDic.Add "jie", -16942
objDic.Add "jin", -16915
objDic.Add "jing", -16733
objDic.Add "jiong", -16708
objDic.Add "jiu", -16706
objDic.Add "ju", -16689
objDic.Add "juan", -16664
objDic.Add "jue", -16657
objDic.Add "jun", -16647
objDic.Add "ka", -16474
objDic.Add "kai", -16470
objDic.Add "kan", -16465
objDic.Add "kang", -16459
objDic.Add "kao", -16452
objDic.Add "ke", -16448
objDic.Add "ken", -16433
objDic.Add "keng", -16429
objDic.Add "kong", -16427
objDic.Add "kou", -16423
objDic.Add "ku", -16419
objDic.Add "kua", -16412
objDic.Add "kuai", -16407
objDic.Add "kuan", -16403
objDic.Add "kuang", -16401
objDic.Add "kui", -16393
objDic.Add "kun", -16220
objDic.Add "kuo", -16216
objDic.Add "la", -16212
objDic.Add "lai", -16205
objDic.Add "lan", -16202
objDic.Add "lang", -16187
objDic.Add "lao", -16180
objDic.Add "le", -16171
objDic.Add "lei", -16169
objDic.Add "leng", -16158
objDic.Add "li", -16155
objDic.Add "lia", -15959
objDic.Add "lian", -15958
objDic.Add "liang", -15944
objDic.Add "liao", -15933
objDic.Add "lie", -15920
objDic.Add "lin", -15915
objDic.Add "ling", -15903
objDic.Add "liu", -15889
objDic.Add "long", -15878
objDic.Add "lou", -15707
objDic.Add "lu", -15701
objDic.Add "lv", -15681
objDic.Add "luan", -15667
objDic.Add "lue", -15661
objDic.Add "lun", -15659
objDic.Add "luo", -15652
objDic.Add "ma", -15640
objDic.Add "mai", -15631
objDic.Add "man", -15625
objDic.Add "mang", -15454
objDic.Add "mao", -15448
objDic.Add "me", -15436
objDic.Add "mei", -15435
objDic.Add "men", -15419
objDic.Add "meng", -15416
objDic.Add "mi", -15408
objDic.Add "mian", -15394
objDic.Add "miao", -15385
objDic.Add "mie", -15377
objDic.Add "min", -15375
objDic.Add "ming", -15369
objDic.Add "miu", -15363
objDic.Add "mo", -15362
objDic.Add "mou", -15183
objDic.Add "mu", -15180
objDic.Add "na", -15165
objDic.Add "nai", -15158
objDic.Add "nan", -15153
objDic.Add "nang", -15150
objDic.Add "nao", -15149
objDic.Add "ne", -15144
objDic.Add "nei", -15143
objDic.Add "nen", -15141
objDic.Add "neng", -15140
objDic.Add "ni", -15139
objDic.Add "nian", -15128
objDic.Add "niang", -15121
objDic.Add "niao", -15119
objDic.Add "nie", -15117
objDic.Add "nin", -15110
objDic.Add "ning", -15109
objDic.Add "niu", -14941
objDic.Add "nong", -14937
objDic.Add "nu", -14933
objDic.Add "nv", -14930
objDic.Add "nuan", -14929
objDic.Add "nue", -14928
objDic.Add "nuo", -14926
objDic.Add "o", -14922
objDic.Add "ou", -14921
objDic.Add "pa", -14914
objDic.Add "pai", -14908
objDic.Add "pan", -14902
objDic.Add "pang", -14894
objDic.Add "pao", -14889
objDic.Add "pei", -14882
objDic.Add "pen", -14873
objDic.Add "peng", -14871
objDic.Add "pi", -14857
objDic.Add "pian", -14678
objDic.Add "piao", -14674
objDic.Add "pie", -14670
objDic.Add "pin", -14668
objDic.Add "ping", -14663
objDic.Add "po", -14654
objDic.Add "pu", -14645
objDic.Add "qi", -14630
objDic.Add "qia", -14594
objDic.Add "qian", -14429
objDic.Add "qiang", -14407
objDic.Add "qiao", -14399
objDic.Add "qie", -14384
objDic.Add "qin", -14379
objDic.Add "qing", -14368
objDic.Add "qiong", -14355
objDic.Add "qiu", -14353
objDic.Add "qu", -14345
objDic.Add "quan", -14170
objDic.Add "que", -14159
objDic.Add "qun", -14151
objDic.Add "ran", -14149
objDic.Add "rang", -14145
objDic.Add "rao", -14140
objDic.Add "re", -14137
objDic.Add "ren", -14135
objDic.Add "reng", -14125
objDic.Add "ri", -14123
objDic.Add "rong", -14122
objDic.Add "rou", -14112
objDic.Add "ru", -14109
objDic.Add "ruan", -14099
objDic.Add "rui", -14097
objDic.Add "run", -14094
objDic.Add "ruo", -14092
objDic.Add "sa", -14090
objDic.Add "sai", -14087
objDic.Add "san", -14083
objDic.Add "sang", -13917
objDic.Add "sao", -13914
objDic.Add "se", -13910
objDic.Add "sen", -13907
objDic.Add "seng", -13906
objDic.Add "sha", -13905
objDic.Add "shai", -13896
objDic.Add "shan", -13894
objDic.Add "shang", -13878
objDic.Add "shao", -13870
objDic.Add "she", -13859
objDic.Add "shen", -13847
objDic.Add "sheng", -13831
objDic.Add "shi", -13658
objDic.Add "shou", -13611
objDic.Add "shu", -13601
objDic.Add "shua", -13406
objDic.Add "shuai", -13404
objDic.Add "shuan", -13400
objDic.Add "shuang", -13398
objDic.Add "shui", -13395
objDic.Add "shun", -13391
objDic.Add "shuo", -13387
objDic.Add "si", -13383
objDic.Add "song", -13367
objDic.Add "sou", -13359
objDic.Add "su", -13356
objDic.Add "suan", -13343
objDic.Add "sui", -13340
objDic.Add "sun", -13329
objDic.Add "suo", -13326
objDic.Add "ta", -13318
objDic.Add "tai", -13147
objDic.Add "tan", -13138
objDic.Add "tang", -13120
objDic.Add "tao", -13107
objDic.Add "te", -13096
objDic.Add "teng", -13095
objDic.Add "ti", -13091
objDic.Add "tian", -13076
objDic.Add "tiao", -13068
objDic.Add "tie", -13063
objDic.Add "ting", -13060
objDic.Add "tong", -12888
objDic.Add "tou", -12875
objDic.Add "tu", -12871
objDic.Add "tuan", -12860
objDic.Add "tui", -12858
objDic.Add "tun", -12852
objDic.Add "tuo", -12849
objDic.Add "wa", -12838
objDic.Add "wai", -12831
objDic.Add "wan", -12829
objDic.Add "wang", -12812
objDic.Add "wei", -12802
objDic.Add "wen", -12607
objDic.Add "weng", -12597
objDic.Add "wo", -12594
objDic.Add "wu", -12585
objDic.Add "xi", -12556
objDic.Add "xia", -12359
objDic.Add "xian", -12346
objDic.Add "xiang", -12320
objDic.Add "xiao", -12300
objDic.Add "xie", -12120
objDic.Add "xin", -12099
objDic.Add "xing", -12089
objDic.Add "xiong", -12074
objDic.Add "xiu", -12067
objDic.Add "xu", -12058
objDic.Add "xuan", -12039
objDic.Add "xue", -11867
objDic.Add "xun", -11861
objDic.Add "ya", -11847
objDic.Add "yan", -11831
objDic.Add "yang", -11798
objDic.Add "yao", -11781
objDic.Add "ye", -11604
objDic.Add "yi", -11589
objDic.Add "yin", -11536
objDic.Add "ying", -11358
objDic.Add "yo", -11340
objDic.Add "yong", -11339
objDic.Add "you", -11324
objDic.Add "yu", -11303
objDic.Add "yuan", -11097
objDic.Add "yue", -11077
objDic.Add "yun", -11067
objDic.Add "za", -11055
objDic.Add "zai", -11052
objDic.Add "zan", -11045
objDic.Add "zang", -11041
objDic.Add "zao", -11038
objDic.Add "ze", -11024
objDic.Add "zei", -11020
objDic.Add "zen", -11019
objDic.Add "zeng", -11018
objDic.Add "zha", -11014
objDic.Add "zhai", -10838
objDic.Add "zhan", -10832
objDic.Add "zhang", -10815
objDic.Add "zhao", -10800
objDic.Add "zhe", -10790
objDic.Add "zhen", -10780
objDic.Add "zheng", -10764
objDic.Add "zhi", -10587
objDic.Add "zhong", -10544
objDic.Add "zhou", -10533
objDic.Add "zhu", -10519
objDic.Add "zhua", -10331
objDic.Add "zhuai", -10329
objDic.Add "zhuan", -10328
objDic.Add "zhuang", -10322
objDic.Add "zhui", -10315
objDic.Add "zhun", -10309
objDic.Add "zhuo", -10307
objDic.Add "zi", -10296
objDic.Add "zong", -10281
objDic.Add "zou", -10274
objDic.Add "zu", -10270
objDic.Add "zuan", -10262
objDic.Add "zui", -10260
objDic.Add "zun", -10256
objDic.Add "zuo", -10254
objDic.Add "[chan1]", -6465

Dim strValue : strValue = ""
For i=1 To Len(oValue)
 Dim Str,oNum
 Str=Mid(oValue, i, 1)
  oNum = Asc(Str)
  If oNum>0 And oNum<160 Then
  strValue = strValue & Ucase(Str)
  Else 
  ''If oNum<-20319 Or oNum>-10247 Then
  If oNum<-20319 Or oNum>-5000 Then
  strValue = strValue&""&Str&""
  Else
  arrKeys = objDic.Keys
  arrItems = objDic.Items
  For j=objDic.Count-1 To 0 Step -1
   If arrItems(j)<=oNum Then Exit For
  Next

  If isAll=2 Then
   ''strValue = strValue & UCase(Left(arrKeys(j),1)) & Mid(arrKeys(j),2,10)
   strValue = strValue&Left(arrKeys(j),1)
   
  ElseIf isAll=1 Then
   strValue = strValue&arrKeys(j)
  Else
   strValue = strValue&Left(arrKeys(j),1)
  End If
  
  End If
  
  End If 
   ''strValue = strValue&"" 
    strValue = strValue&""         
Next   
GetSpellingChar = strValue
objDic.RemoveAll
Set objDic = Nothing 
End Function   

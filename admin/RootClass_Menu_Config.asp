<%
'�˵���ʾ��������
Const RCM_Menu_1="4"      '�˵�������ʽ 1����  2����  3����  4����
Const RCM_Menu_2="0"      '�˵���������ƫ����
Const RCM_Menu_3="0"      '�˵���������ƫ����
Const RCM_Menu_4="2"      '�˵���߾�
Const RCM_Menu_5="3"      '�˵�����
Const RCM_Menu_6="6"      '�˵�����߾�
Const RCM_Menu_7="7"      '�˵����ұ߾�
Const RCM_Menu_8="100"      '�˵�͸����         0-100 ��ȫ͸��-��ȫ��͸��
Const RCM_Menu_9="filter:Glow(Color=#000000, Strength=3)"      '������Ч
Const RCM_Menu_10="4"        '���ָ�ڲ˵���ʱ���˵�����Ч��
Const RCM_Menu_11=""        '������Ч
Const RCM_Menu_12="23"        '����Ƴ��˵���ʱ���˵�����Ч��
Const RCM_Menu_13="50"        '�˵�����Ч���ٶ�  10-100
Const RCM_Menu_14="2"        '�����˵���ӰЧ�� 0��none  1��simple  2��complex
Const RCM_Menu_15="4"        '�����˵���Ӱ���
Const RCM_Menu_16="#999999"        '�����˵���Ӱ��ɫ
Const RCM_Menu_17="#ffffff"        '�����˵�������ɫ
Const RCM_Menu_18=""        '�����˵�����ͼƬ��ֻ�е��˵������ɫ��Ϊ͸��ɫ��transparent ʱ����Ч
Const RCM_Menu_19="3"        '�����˵�����ͼƬƽ��ģʽ�� 0����ƽ��  1������ƽ��  2������ƽ��  3����ȫƽ��
Const RCM_Menu_20="1"        '�����˵��߿����� 0���ޱ߿�  1����ʵ��  2��˫ʵ��  5������  6��͹��
Const RCM_Menu_21="1"        '�����˵��߿����
Const RCM_Menu_22="#ACA899"        '�����˵��߿���ɫ
Const RCM_Menu_23="#ffffff"

'�˵����������
Const RCM_Item_1="0"      '�˵�������  0--Txt  1--Html  2--Image
Const RCM_Item_2=""       '�˵�������
Const RCM_Item_3=""       '�˵���ΪImage��ͼƬ�ļ�
Const RCM_Item_4=""       '�˵���ΪImage�����ָ�ڲ˵���ʱ��ͼƬ�ļ���
Const RCM_Item_5="-1"     '�˵���ΪImage��ͼƬ����
Const RCM_Item_6="-1"     '�˵���ΪImage��ͼƬ�߶�
Const RCM_Item_7="0"      '�˵���ΪImage��ͼƬ�߿�
Const RCM_Item_8=""       '�˵������ӵ�ַ
Const RCM_Item_9=""       '�˵�������Ŀ�� �磺_self  _blank
Const RCM_Item_10=""      '�˵�������״̬����ʾ
Const RCM_Item_11=""      '�˵������ӵ�ַ��ʾ��Ϣ
Const RCM_Item_12=""        '�˵�����ͼƬ
Const RCM_Item_13=""        '���ָ�ڲ˵���ʱ���˵�����ͼƬ
Const RCM_Item_14="0"        '�˵�����ͼƬ���ȣ�0Ϊͼ���ļ�ԭʼֵ
Const RCM_Item_15="0"        '�˵�����ͼƬ�߶ȣ�0Ϊͼ���ļ�ԭʼֵ
Const RCM_Item_16="0"        '�˵�����ͼƬ�߿��С
Const RCM_Item_17=""        '�˵�����ͼƬ���磺arrow_r.gif
Const RCM_Item_18=""        '���ָ�ڲ˵���ʱ���˵�����ͼƬ���磺arrow_w.gif
Const RCM_Item_19="0"        '�˵�����ͼƬ���ȣ�0Ϊͼ���ļ�ԭʼֵ
Const RCM_Item_20="0"        '�˵�����ͼƬ�߶ȣ�0Ϊͼ���ļ�ԭʼֵ
Const RCM_Item_21="0"        '�˵�����ͼƬ�߿��С
Const RCM_Item_22="0"        '�˵�������ˮƽ���뷽ʽ  0�������  1������  2���Ҷ���
Const RCM_Item_23="1"        '�˵������ִ�ֱ���뷽ʽ  0������  1������  2���ײ�
Const RCM_Item_24="#F1F2EE"        '�˵������ɫ  ͸��ɫ��'transparent'
Const RCM_Item_25="1"        '�˵������ɫ�Ƿ���ʾ  0����ʾ  ����������ʾ
Const RCM_Item_26="#CCCCCC"        '���ָ�ڲ˵���ʱ���˵������ɫ
Const RCM_Item_27="1"        '���ָ�ڲ˵���ʱ���˵������ɫ�Ƿ���ʾ��  0����ʾ  ����������ʾ
Const RCM_Item_28=""        '�˵����ͼƬ
Const RCM_Item_29=""        '���ָ�ڲ˵���ʱ���˵����ͼƬ
Const RCM_Item_30="3"        '�˵����ͼƬƽ��ģʽ�� 0����ƽ��  1������ƽ��  2������ƽ��  3����ȫƽ��
Const RCM_Item_31="3"     '���ָ�ڲ˵���ʱ���˵����ͼƬƽ��ģʽ��0-3
Const RCM_Item_32="0"        '�˵���߿����� 0���ޱ߿�  1����ʵ��  2��˫ʵ��  5������  6��͹��
Const RCM_Item_33="0"        '�˵���߿����
Const RCM_Item_34="#FFFFF7"        '�˵���߿���ɫ
Const RCM_Item_35="#FF0000"        '���ָ�ڲ˵���ʱ���˵���߿���ɫ
Const RCM_Item_36="#000000"        '�˵���������ɫ
Const RCM_Item_37="#CC0000"        '���ָ�ڲ˵���ʱ���˵���������ɫ
Const FontSize_RCM_Item_38="9pt"        '�˵������ִ�С
Const FontName_RCM_Item_38="����"        '�˵�����������
Const FontSize_RCM_Item_39="9pt"        '���ָ�ڲ˵���ʱ,�˵������ִ�С
Const FontName_RCM_Item_39="����"        '���ָ�ڲ˵���ʱ,�˵�����������
%>
Attribute VB_Name = "databuf"
'                  ͼ�λ������淶��
'һ , ˵��:
'      ���ڿ��Ƴ�����ͨѶ�ӳ��򣻿��Ƴ��������г��򣬱�������֮��Ĺ�ϵ ��
'�� , �淶:
'ȫ�ֱ���:
'Global Const CHINA = True
'Global Const ENGLISH = False
'Global china_english As Boolean
'Global Run_flag As Integer

Type elec_in_table
            name                   As String * 10                 '���ź���
            str                        As String * 7                 '����--->����
            run        As Integer                   '�𶯽ӵ�
            run_set As Integer                  '�߼��趨ֵ
                 
            check As Integer                '���ӵ�
            check_set As Integer        '�߼��趨ֵ
            run_value As Integer
            check_value As Integer
            check_time        As Integer               '���ʱ��
            flag         As Integer     'ͣ���־ (����)
            Value      As Integer
            count      As Integer            '���� (����)
            error_name   As String * 10    '�������
End Type
 '               �� {"����˨","��",1001,1,1002,0,0,1,10,0,0,"111��}


'      ���źŻ�����(read only)            ÿ0.5Sһ��
'        Type elect_send_type
 '           elec_in     As Integer                      '=1  ����д   =0  �ɶ�
  '          Length      As Integer                      'ʵ�ʳ���
   '         Note_Name(100)   As Integer   '�ڵ���
    '        data(100)   As Integer          '����
  '      End Type
  
'        Global elect_receive_input       As elect_send_type         '�����ڵ�
'        Global elect_receive_output       As elect_send_type        '���ڵ�

'         ���ݻ�����(timing  read only)

'             If duan_flag <> 2 Then
'                    Exit Sub
'             End If
'            ÿ����һ��
            Type make_lian_liao
                  duan_write      As Integer            ' д�α�־  ��0 �ɶ�   ��1 ��д    ���� �м����
                  duan_hao       As Integer
                  set_hun_time   As Integer          '/* �趨����ʱ�� */
                  get_hun_time   As Integer          '/* ʵ�ʻ���ʱ�� */
                  set_you_time   As Integer             '/* ��Ͷ��    Ϊ0 ��Ͷ*/
                  set_tan_time   As Integer        '/* ��Ͷ̿   Ϊ 0 ��Ͷ*/
                  set_jiao_time  As Integer        '/* ��Ͷ��  Ϊ  0 ��Ͷ*/
                  set_xiao_time  As Integer       ' /* ��ͶС��  Ϊ 0 ��Ͷ*/
                  get_you_time   As Integer        '/* ��Ͷ��    Ϊ0 ��Ͷ*/
                  get_tan_time   As Integer       '/* ��Ͷ̿   Ϊ 0 ��Ͷ*/
                  get_jiao_time  As Integer       '/* ��Ͷ��  Ϊ  0 ��Ͷ*/
                  get_xiao_time As Integer      ' /* ��ͶС��  Ϊ 0 ��Ͷ*/

                  next_tempro     As Integer
                  next_power      As Integer

            End Type
            Global make_input  As make_lian_liao

            'ÿ������һ��
            'If che_flag <> 2 Then
             '   exit
            'End If

            Type make_pei_fan
                   pei_number      As String * 6
                   give_time    As Integer
                   duan_hao     As Integer
                   cai_number   As String * 6
                   cai_name     As String * 15
                   mathine     As Integer          '/*�����    0    1 2 */
                   wna         As String * 5            '/* ����  ban1       ̿����1*/
                                                                                            'ban2 ̿����2
                                                                                            'ban3 ���ϳ�
                                                                                            'ban4 ���ϳ�
                    dou           As String * 5        '/* ����     dou1
                                                                            '��1 dou2
                                                                            '��2 dou3
                   set_weight   As Integer            '/* ��׼���� */
                   set_gon_cha  As Integer           ' /* �趨��� */
                   get_weight   As Integer            '/* ʵ������ */
            End Type
'            eg:   1  "a1" 0,"1","rr",1,̿����2,��5,10,10,39
'            eg:   1  "a1" 0,"1","rr",1,̿����2,��6,10,10,12
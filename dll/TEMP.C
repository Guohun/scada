#include <time.h>
#include <sys/types.h>
#include <sys/timeb.h>
  #include "windows.h"
  #include "SQL.H"
  #include "SQLext.H"
  #include <string.h> 
  #include "dll.h"
  #include "temp.h"
             

//每段写一次
//全、局变量    duan_begin_time   ---->c1.c   write_power
int  WINAPI   write_make(struct make_lian_liao temp)
{
    char  select_str[400];
    HENV henv;
	HDBC hdbc; 
	HSTMT hstmt; 
	RETCODE retcode;
    int   _duan;
    struct make_lian_liao  get_make_lian_liao={0,0};

        _duan=temp.duan_hao;

        get_make_lian_liao.only_key=get_make_table.only_key;
        get_make_lian_liao.now_che=get_make_table.now_che;
       //get_make_lian_liao.mathine=get_make_table.mathine;
        get_make_lian_liao.duan_hao=get_make_table.now_duan;

        get_make_lian_liao.set_hun_time=send_pei_fan.huan[_duan].lian_time;
        get_make_lian_liao.set_you_time=send_pei_fan.huan[_duan].dou_you_time;
        get_make_lian_liao.set_tan_time=send_pei_fan.huan[_duan].dou_tan_time;
        get_make_lian_liao.set_jiao_time=send_pei_fan.huan[_duan].dou_jiao_time;
        get_make_lian_liao.set_xiao_time=send_pei_fan.huan[_duan].dou_xiao_time;

        get_make_lian_liao.get_hun_time=temp.get_hun_time;
        get_make_lian_liao.get_you_time=temp.get_you_time;
        get_make_lian_liao.get_tan_time=temp.get_tan_time;
        get_make_lian_liao.get_jiao_time=temp.get_jiao_time;
        get_make_lian_liao.get_xiao_time=temp.get_xiao_time;

        lstrcpy(get_make_lian_liao.duan_begin_time,duan_begin_time);

	retcode = SQLAllocEnv(&henv); /* Environment handle */ 
	if (retcode == SQL_SUCCESS) 
	{ 
        retcode = SQLAllocConnect(henv, &hdbc); /* Connection handle */
        if (retcode == SQL_SUCCESS)
        {
            /* Set login timeout to 5 seconds. */
            SQLSetConnectOption(hdbc, SQL_LOGIN_TIMEOUT, 5);
            /* Connect to data source */
            retcode = SQLConnect(hdbc, (unsigned char *) "ljxt", SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
            if (retcode == SQL_SUCCESS || retcode == SQL_SUCCESS_WITH_INFO)
            {
                /* Process data after successful connection */
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert make_lian_liao \
set (only_key,now_che,duan_hao, set_huan_time,get_huan_time,set_you_time,\
set_tan_time,set_jiao_time,set_xiao_time,get_you_time,get_tan_time,get_jiao_time,\
next_termpro,next_power,duan_begin_time,duan_begin_time,duan_end_tim,\
power,tempro)set = (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); //total 20 ?

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
        //now_che       SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, name, 0, NULL);
                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(_duan), 0,NULL );
                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(send_pei_fan.huan[_duan].lian_time),0,NULL);
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.get_hun_time), 0,NULL );
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(send_pei_fan.huan[_duan].dou_you_time),0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(send_pei_fan.huan[_duan].dou_jiao_time),0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(send_pei_fan.huan[_duan].dou_tan_time),0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(send_pei_fan.huan[_duan].dou_xiao_time),0,NULL);

                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.get_you_time), 0,NULL );
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.get_tan_time), 0,NULL );
                        SQLBindParameter(hstmt, 12, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.get_jiao_time), 0,NULL );
                        SQLBindParameter(hstmt, 13, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.get_xiao_time), 0,NULL );

                        SQLBindParameter(hstmt, 14, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.next_tempro), 0,NULL );
                        SQLBindParameter(hstmt, 15, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.next_power), 0,NULL );
                        SQLBindParameter(hstmt, 16, SQL_PARAM_INPUT, SQL_C_DATE, SQL_DATE, 0, 0, &(get_make_lian_liao.duan_begin_time), 0,NULL );

                        _strtime( duan_begin_time );
                        SQLBindParameter(hstmt, 17, SQL_PARAM_INPUT, SQL_C_TIME, SQL_TIME, 0, 0, &duan_begin_time, 0,NULL );

                        SQLBindParameter(hstmt, 18, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,300, 0, work_power.power, 0,NULL );
                        SQLBindParameter(hstmt, 19, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,work_power.total, 0, work_power.tempro, 0,NULL );



                        retcode = SQLExecute(hstmt);         /* Execute statement with */
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

          SQLDisconnect(hdbc);
        }// end if SQLConnect
       SQLFreeConnect(hdbc);
       }//end if allocconnect
      SQLFreeEnv(henv);
    }//end if allocenv
	return 1;
}

int WINAPI   write_che(struct make_base_type temp)
{
    char  select_str[400];
    HENV henv;
	HDBC hdbc; 
	HSTMT hstmt; 
	RETCODE retcode;
    char  tmpbuf[128];

    get_make_table.only_key++;// next che

	retcode = SQLAllocEnv(&henv); /* Environment handle */ 
	if (retcode == SQL_SUCCESS) 
	{ 
        retcode = SQLAllocConnect(henv, &hdbc); /* Connection handle */
        if (retcode == SQL_SUCCESS)
        {
            /* Set login timeout to 5 seconds. */
            SQLSetConnectOption(hdbc, SQL_LOGIN_TIMEOUT, 5);
            /* Connect to data source */
            retcode = SQLConnect(hdbc, (unsigned char *) "ljxt", SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
            if (retcode == SQL_SUCCESS || retcode == SQL_SUCCESS_WITH_INFO)
            {
                int i=0;
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
					    SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.th1[i].mathine, 0, NULL);
//                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL ); //duan_hao

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.th1[i].add_time,0,NULL);  //give_time
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.th1[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.th1[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.th1[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.th1[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.th1[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.th1[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.th1[i].gon_cha),0,NULL);
                        for(i=0;i<temp.th1_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.th2[i].mathine, 0, NULL);
//                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL ); //duan_hao

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.th2[i].add_time,0,NULL);  //give_time
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.th2[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.th2[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.th2[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.th2[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.th2[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.th2[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.th2[i].gon_cha),0,NULL);

                        for(i=0;i<temp.th2_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

// insert into sj1
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.sj1[i].mathine, 0, NULL);
//                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL ); //duan_hao

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.sj1[i].add_time,0,NULL);  //give_time
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.sj1[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.sj1[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.sj1[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.sj1[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.sj1[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.sj1[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.sj1[i].gon_cha),0,NULL);

                        for(i=0;i<temp.sj1_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);
// insert into sj2
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.sj2[i].mathine, 0, NULL);
//                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL ); //duan_hao

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.sj2[i].add_time,0,NULL);  //give_time
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.sj2[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.sj2[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.sj2[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_CHAR,5, 0, temp.sj2[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.sj2[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.sj2[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.sj2[i].gon_cha),0,NULL);
                        for(i=0;i<temp.sj2_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

// insert into yz1
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {

   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao
                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.yz1[i].mathine, 0, NULL);
             //         SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.duan_hao, 0,NULL );

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yz1[i].add_time,0,NULL);
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yz1[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yz1[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yz1[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yz1[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yz1[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yz1[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yz1[i].gon_cha),0,NULL);
                        for(i=0;i<temp.yz1_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);
// insert into yz2
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao
                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.yz1[i].mathine, 0, NULL);
//                      SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL );

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yz2[i].add_time,0,NULL);
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yz2[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yz2[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yz2[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yz2[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yz2[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yz2[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yz2[i].gon_cha),0,NULL);
                        for(i=0;i<temp.yz2_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

// insert into yt1
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {

   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.yt1[i].mathine, 0, NULL);
//                      SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL );

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yt1[i].add_time,0,NULL);
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yt1[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yt1[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yt1[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yt1[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yt1[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yt1[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yt1[i].gon_cha),0,NULL);
                        for(i=0;i<temp.yt1_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);
// insert into yt2
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {

   lstrcpy(select_str," insert make_pei_fan \
set (only_key,mathine,duan_hao,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha)set = (?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao
                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 15, 0, temp.yt1[i].mathine, 0, NULL);
//                      SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL );

                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yt2[i].add_time,0,NULL);
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yt2[i].cai_number,0,NULL);
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yt2[i].name,0,NULL);
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yt2[i].wna,0,NULL);
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,5, 0, temp.yt2[i].dou,0,NULL);
                        SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yt2[i].weight),0,NULL);
                        SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yt2[i].get_weight),0,NULL);
                        SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT,0, 0, &(temp.yt2[i].gon_cha),0,NULL);
                        for(i=0;i<temp.yt2_sum;i++)
                        {
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

// insert into make_base
/*
            get_make_base.only_key=get_make_table.only_key;
            get_make_base.now_che=get_make_table.now_che;
            get_make_base.set_weight=send_pei_fan.base.weight;
//          get_make_base.name
            get_make_base.set_pai_time=send_pei_fan.base.pai_liao_time;
*/
                retcode = SQLAllocStmt(hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert make_base \
   set (only_key,now_che,set_weight,get_weight,\
set_pai_time,get_pai_time,huan_time,pai_tempro)set = (?,?,?,?,?,?,?,?,?) "); //total 9 ?  duan_hao // not name ,all_power goncha

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &get_make_table.now_che, 0,NULL );
                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &send_pei_fan.base.weight, 0,NULL );
                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.get_weight, 0,NULL );
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &send_pei_fan.base.pai_liao_time, 0,NULL );
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.get_pai_time, 0,NULL );
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.huan_time, 0,NULL );
                        SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.pai_tempro, 0,NULL );
                        retcode = SQLExecute(hstmt);         /* Execute statement with */
                        
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);




          SQLDisconnect(hdbc);
        }// end if SQLConnect
       SQLFreeConnect(hdbc);
       }//end if allocconnect
      SQLFreeEnv(henv);
    }//end if allocenv
	return 1;
}
/*

int  main()
{
	int i;
	 struct Produce_type temp;
	can_read("12","test",12);
	read_produce(&temp);

    printf("max=%d min=%d\n",temp.base.min_wen_du,
			temp.base.max_wen_du);
	printf("duan  dou_you_time,dou_tan_time,st,speed");
    for(i=0;i<temp.pei_max;i++)
        printf("%6d %6d %8d %3d %4d",
            temp.huan.duan,temp.huan.dou_you_time,
            temp.huan.dou_tan_time,temp.huan.st,
            temp.huan.speed);
}
*/

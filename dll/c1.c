  /* sql_dll.cpp
  */
#include <time.h>
#include <sys/types.h>
#include <sys/timeb.h>

  #include <windows.h>
  #include "SQL.H"
  #include "SQLext.H"
  #include <string.h>
  #include "dll.h"
#include "temp.h"
/* 1997.12.6 朱国魂 修改:
	  现象：胶料没选
 	NOMANGLE int CCONV  can_read(BSTR * pei_number,BSTR *name,int ban,char total_che,int flag)	
 int ---->short
	//jiao_liao_table get add_time=1,or 2 sj1  get away
	
*/
/* 1998.2.6 朱国魂 修改:
	  现象：传递配方有误
 	NOMANGLE int CCONV  can_read(BSTR * pei_number,BSTR *name,int ban,char total_che,int flag)	
	less mathine
*/


//move to dll.c
#pragma     data_seg("sdata")
	int  ljxt_mdb_open_flag=0;
	HENV ljxt_henv=NULL;
	HDBC ljxt_hdbc=NULL; 
	extern int  Prev_Run_Flag; 	
#pragma data_seg()



  
/*
	flag 本用于  simulate or  run
          现用于  simulate or run <<8|mathine
                or  在线修改  3
*/
 NOMANGLE int CCONV  can_read(BSTR * pei_number,BSTR *name,short ban,char total_che,short flag)
  //由VB-->DLL 送入  传递运行配方
  //可能重复送
  /*  vb--->send_pei_fan
  */
   {
	HSTMT hstmt; 
	RETCODE retcode;
 	char select_str[580];

	control_can_read =CAN_READ-1;   //不可读
	
	lstrcpy(send_run_table.name,name);
	lstrcpy(send_run_table._number,pei_number);
	send_run_table.total_che= total_che;
	if((flag>>4)!=3)
	{
		send_run_table.now_che=0;
		send_run_table.all_time=0;	  /*总工作时间 */
		send_run_table.all_che_time=0;	/*总一车时间 */		
		send_run_table.duan_time=0; /*段时间 */			
	}
		send_run_table.all_duan=0;	
		send_run_table.ban=ban;	/*班  */
		send_run_table.mathine=flag&0x3;  /*机号 */

if (ljxt_mdb_open_flag==0)
{
	ljxt_mdb_open_flag=1;
	retcode = SQLAllocEnv(&ljxt_henv); /* Environment handle */ 
	if (retcode == SQL_SUCCESS) 
	{ 
        retcode = SQLAllocConnect(ljxt_henv, &ljxt_hdbc); /* Connection handle */
        if (retcode == SQL_SUCCESS)
        {
            SQLSetConnectOption(ljxt_hdbc, SQL_LOGIN_TIMEOUT, 5);
            retcode = SQLConnect(ljxt_hdbc, (unsigned char *) "ljxt", SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
			if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
			{
				write_error(ODBC_ERROR,"没设置ljxt ODBC");				
				SQLFreeEnv(ljxt_henv);
				ljxt_henv=NULL;
				ljxt_mdb_open_flag=0;
				return FALSE;
			}			
		}//end SQLALLOCConnect
	}//end SQLALLocEnv
};//end if comm_mdb_open_flag


	
				/* Process data after successful connection */
                retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
lstrcpy(select_str,"select duan, dou_you_time,\
dou_tan_time,dou_jiao_time,dou_xiao_time,st,speed,pre,\
lian_time,upper_pre,tem,neg,ctr,stop_time,Dou_you2_time \
from  lain_liao_table where pei_number like \'");							
                            lstrcat(select_str,pei_number);							
							if (send_run_table.mathine==2)
									lstrcat(select_str,"\' and mathine=2");														
							else
									lstrcat(select_str,"\' and mathine=1");														
                            lstrcat(select_str," order by duan");

                    retcode=lstrlen(select_str);
//                  //printf("len=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        int i=0;
                       while (TRUE)
                        {
						   DWORD   temp_int;
                        SQLBindCol(hstmt,1,SQL_C_SLONG,&send_pei_fan.huan[i].duan,0,&temp_int);
                        SQLBindCol(hstmt,2,SQL_C_SLONG,&send_pei_fan.huan[i].dou_you_time,0,&temp_int);
                        SQLBindCol(hstmt,3,SQL_C_SLONG,&send_pei_fan.huan[i].dou_tan_time,0,&temp_int);
                        SQLBindCol(hstmt,4,SQL_C_SLONG,&send_pei_fan.huan[i].dou_jiao_time,0,&temp_int);
                        SQLBindCol(hstmt,5,SQL_C_SLONG,&send_pei_fan.huan[i].dou_xiao_time,0,&temp_int);
                        SQLBindCol(hstmt,6,SQL_C_SLONG,&send_pei_fan.huan[i].st,0,&temp_int);
                        SQLBindCol(hstmt,7,SQL_C_STINYINT,&send_pei_fan.huan[i].speed,0,&temp_int);
                        SQLBindCol(hstmt,8,SQL_C_STINYINT,&send_pei_fan.huan[i].pre,0,&temp_int);
                        SQLBindCol(hstmt,9,SQL_C_STINYINT,&send_pei_fan.huan[i].lian_time,0,&temp_int);
                        SQLBindCol(hstmt,10,SQL_C_STINYINT,&send_pei_fan.huan[i].upper_pre,0,&temp_int);

                        SQLBindCol(hstmt,11,SQL_C_SLONG,&send_pei_fan.huan[i].tem,0,&temp_int);
                        SQLBindCol(hstmt,12,SQL_C_FLOAT,&send_pei_fan.huan[i].neg,0,&temp_int);

                        SQLBindCol(hstmt,13,SQL_C_STINYINT,&send_pei_fan.huan[i].ctr,0,&temp_int);
                        SQLBindCol(hstmt,14,SQL_C_SLONG,&send_pei_fan.huan[i].stop_time,0,&temp_int);
                        SQLBindCol(hstmt,15,SQL_C_SLONG,&send_pei_fan.huan[i].dou_you2_time,0,&temp_int);
                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*                            printf("i= %d record=%-5d %-2d %6d  %4d\n",i,send_pei_fan.huan[i].duan,
                                    send_pei_fan.huan[i].dou_you_time, 
                                    send_pei_fan.huan[i].dou_tan_time,
                                    send_pei_fan.huan[i].dou_jiao_time);
*/
                            i++;
                          }
                          send_pei_fan.pei_max=i ;  //from 0 begin
					}//end if sqlexec
					SQLTransact(ljxt_henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
                    SQLFreeStmt(hstmt, SQL_DROP);


/*search  lian_add_type base */
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
 lstrcpy(select_str,"select  min_wen_du,\
  max_wen_du,dou_liao_time,dou_liao_start,\
 weigh,pai_liao_time,pai_liao_ask,temp1,\
temp2,temp3 from  lian_add_table where pei_number like \'");
                            lstrcat(select_str,pei_number);
                            if (send_run_table.mathine==2)
                                  lstrcat(select_str,"\' and mathine=2");                                                                                                  
                            else
                                  lstrcat(select_str,"\' and mathine=1");                                                                                                  
                    retcode=lstrlen(select_str);
                    //printf("\nlen=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
						   DWORD   temp_int;
						SQLBindCol(hstmt,1,SQL_C_SLONG,&send_pei_fan.base.min_wen_du,0,&temp_int);
                        SQLBindCol(hstmt,2,SQL_C_SLONG,&send_pei_fan.base.max_wen_du,0,&temp_int);
  //                      SQLBindCol(hstmt,3,SQL_C_SLONG,&send_pei_fan.base.dou_liao_time,0,&temp_int);
//                        SQLBindCol(hstmt,4,SQL_C_SLONG,&send_pei_fan.base.dou_liao_start,0,&temp_int);
                        SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.base.weight,0,&temp_int);
                        SQLBindCol(hstmt,6,SQL_C_SLONG,&send_pei_fan.base.pai_liao_time,0,&temp_int);
                        SQLBindCol(hstmt,7,SQL_C_SLONG,&send_pei_fan.base.pai_liao_ask,0,&temp_int);
                        SQLBindCol(hstmt,8,SQL_C_SLONG,&send_pei_fan.base.temp1,0,&temp_int);

                        SQLBindCol(hstmt,9,SQL_C_FLOAT,&send_pei_fan.base.temp2,0,&temp_int);
                        SQLBindCol(hstmt,10,SQL_C_SLONG,&send_pei_fan.base.temp3,0,&temp_int);
                            retcode = SQLFetch(hstmt);
                    }//end if sqlexec
					SQLTransact(ljxt_henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);




 // search  for  th1
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
                    int i=0;

                    lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.dou,a.fast_do,a.drop_do,a.control_time,a.add_time,a.stop_time,\
a.gon_cha,a.weight,b.mathine from  tan_hei_table a,cai_liao_table b where  add_time=1 and  a.cai_number=b.cai_number and a.mathine=b.mathine  and a.pei_number like \'");
                            lstrcat(select_str,pei_number);							
                            lstrcat(select_str,"\' order by a.weight desc");

                    retcode=lstrlen(select_str);
                    //printf("\nlen=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                     DWORD  temp;
                     while (TRUE)
                     {
                            SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.th1[i].cai_number,6,&temp);
                            SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.th1[i].name,15,&temp);
                            SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.th1[i].wna,5,&temp);
                            SQLBindCol(hstmt,4,SQL_C_CHAR,send_pei_fan.th1[i].dou,5,&temp);

                            SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.th1[i].fast_do,0,NULL);
                            SQLBindCol(hstmt,6,SQL_C_FLOAT,&send_pei_fan.th1[i].drop_do,0,NULL);
                            SQLBindCol(hstmt,7,SQL_C_FLOAT,&send_pei_fan.th1[i].control_time,0,NULL);
                            SQLBindCol(hstmt,8,SQL_C_SLONG,&send_pei_fan.th1[i].add_time,0,NULL);
                            SQLBindCol(hstmt,9,SQL_C_SLONG,&send_pei_fan.th1[i].stop_time,0,NULL);

                            SQLBindCol(hstmt,10,SQL_C_FLOAT,&send_pei_fan.th1[i].gon_cha,0,NULL);
                            SQLBindCol(hstmt,11,SQL_C_FLOAT,&send_pei_fan.th1[i].weight,0,NULL);
                            SQLBindCol(hstmt,12,SQL_C_SLONG,&send_pei_fan.th1[i].mathine,0,NULL);

                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*                          //printf("%s %s %s\n",send_pei_fan.th1[i].name,
                                send_pei_fan.th1[i].wna, send_pei_fan.th1[i].dou);
*/
							if (send_run_table.mathine==send_pei_fan.th1[i].mathine)
									i++;
                          }//END WHILE
                           send_pei_fan.th1_sum=i;
					}//end if sqlexec
//                    SQLTransact(henv, hdbc, SQL_COMMIT);
					SQLTransact(ljxt_henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

 // search  for  th2
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
                    int i=0;

                    lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.dou,a.fast_do,a.drop_do,a.control_time,a.add_time,a.stop_time,\
a.gon_cha,a.weight,b.mathine from  tan_hei_table a,cai_liao_table b where  add_time=2 and  a.cai_number=b.cai_number and a.mathine=b.mathine and a.pei_number like \'");
                            lstrcat(select_str,pei_number);
                            lstrcat(select_str,"\' order by a.weight desc");

                    retcode=lstrlen(select_str);
                    //printf("\nlen=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        DWORD  temp;
						while (TRUE)
                         {
                      
                        SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.th2[i].cai_number,5,&temp);
                        SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.th2[i].name,14,&temp);
                        SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.th2[i].wna,4,&temp);
                        SQLBindCol(hstmt,4,SQL_C_CHAR,send_pei_fan.th2[i].dou,4,&temp);

                        SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.th2[i].fast_do,0,NULL);
                        SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.th2[i].fast_do,0,NULL);
                        SQLBindCol(hstmt,6,SQL_C_FLOAT,&send_pei_fan.th2[i].drop_do,0,NULL);
                        SQLBindCol(hstmt,7,SQL_C_FLOAT,&send_pei_fan.th2[i].control_time,0,NULL);
                        SQLBindCol(hstmt,8,SQL_C_SLONG,&send_pei_fan.th2[i].add_time,0,NULL);
                        SQLBindCol(hstmt,9,SQL_C_SLONG,&send_pei_fan.th2[i].stop_time,0,NULL);

                            SQLBindCol(hstmt,10,SQL_C_FLOAT,&send_pei_fan.th2[i].gon_cha,0,NULL);
                            SQLBindCol(hstmt,11,SQL_C_FLOAT,&send_pei_fan.th2[i].weight,0,NULL);
                            SQLBindCol(hstmt,12,SQL_C_SLONG,&send_pei_fan.th2[i].mathine,0,NULL);


                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
                            if (retcode == SQL_ERROR )
                            {
                                            //show_error();
                                            break;
                            }
/*                             //printf("%-5s %5s %5s\n",send_pei_fan.th2[i].name,
                                send_pei_fan.th2[i].wna, send_pei_fan.th2[i].dou);
*/
                        if (send_run_table.mathine==send_pei_fan.th2[i].mathine)
                            i++;
                          }
                           send_pei_fan.th2_sum=i;
					}//end if sqlexec
                    SQLTransact(ljxt_henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);


 // search  for  sj1
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {

                    int i=0;
lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.jiao,a.control_time,a.add_time,a.stop_time,\
a.weight,a.gon_cha,b.mathine,a.sort_idx from  jiao_liao_table a,cai_liao_table b where  a.add_time=1 and  a.cai_number=b.cai_number and a.mathine=b.mathine  and a.pei_number like \'");
                            lstrcat(select_str,pei_number);
                            lstrcat(select_str,"\' order by a.sort_idx");
/*
                    lstrcpy(select_str,"select  cai_number,\
                        name,wna,st,control_time,add_time,\
                stop_time, weight,gon_cha  from  jiao_liao_table where  add_time=1 and  pei_number like \'");
*/

//                    retcode=lstrlen(select_str);
                    //printf("\nlen=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                         while (TRUE)
                         {
                        SDWORD  temp;
                        SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.sj1[i].cai_number,6,&temp);
                        SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.sj1[i].name,15,&temp);
                        SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.sj1[i].wna,4,&temp);
                        SQLBindCol(hstmt,4,SQL_C_SLONG,&send_pei_fan.sj1[i].stop_time,0,&temp);
//                        SQLBindCol(hstmt,5,SQL_C_SLONG,&send_pei_fan.sj1[i].control_time,0,NULL);
                        SQLBindCol(hstmt,6,SQL_C_SLONG,&send_pei_fan.sj1[i].add_time,0,&temp);
//                        SQLBindCol(hstmt,7,SQL_C_SLONG,&send_pei_fan.sj1[i].stop_time,0,NULL);

                        SQLBindCol(hstmt,8,SQL_C_FLOAT,&send_pei_fan.sj1[i].weight,0,&temp);
                        SQLBindCol(hstmt,9,SQL_C_FLOAT,&send_pei_fan.sj1[i].gon_cha,0,&temp);
                        SQLBindCol(hstmt,10,SQL_C_SLONG,&send_pei_fan.sj1[i].mathine,0,&temp);


                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*                          //printf("%-5s %5s %5d\n",send_pei_fan.sj1[i].name,
                                send_pei_fan.sj1[i].wna, send_pei_fan.sj1[i].weight);
*/
//                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
//                                    break;      //exit  while
                        if (send_run_table.mathine==send_pei_fan.sj1[i].mathine)
                            i++;


                          }
                           send_pei_fan.sj1_sum=i;
					}//end if sqlexec
//                    SQLTransact(henv, hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);



 // search  for  sj2
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
                    int i=0;
                    lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.jiao,a.control_time,a.add_time,a.stop_time,\
a.weight,a.gon_cha,b.mathine,a.sort_idx from  jiao_liao_table a,cai_liao_table b where  add_time=2 and  a.cai_number=b.cai_number and a.mathine=b.mathine and a.pei_number like \'");
                            lstrcat(select_str,pei_number);
                            lstrcat(select_str,"\' order by a.sort_idx");
                    retcode=lstrlen(select_str);
                    //printf("\nlen=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                         while (TRUE)
                         {
							 DWORD  temp;
							SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.sj2[i].cai_number,6,&temp);
							SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.sj2[i].name,15,&temp);
                            SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.sj2[i].wna,8,&temp);
                            SQLBindCol(hstmt,4,SQL_C_SLONG,&send_pei_fan.sj2[i].stop_time,0,&temp);//st
//		                    SQLBindCol(hstmt,5,SQL_C_SLONG,&send_pei_fan.sj2[i].control_time,0,NULL);
							SQLBindCol(hstmt,6,SQL_C_SLONG,&send_pei_fan.sj2[i].add_time,0,&temp);
							//SQLBindCol(hstmt,7,SQL_C_SLONG,&send_pei_fan.sj2[i].stop_time,0,NULL);

							SQLBindCol(hstmt,8,SQL_C_FLOAT,&send_pei_fan.sj2[i].weight,0,&temp);
							SQLBindCol(hstmt,9,SQL_C_FLOAT,&send_pei_fan.sj2[i].gon_cha,0,&temp);
							SQLBindCol(hstmt,10,SQL_C_SLONG,&send_pei_fan.sj2[i].mathine,0,&temp);

                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*							//printf("%-5s %5s %5d\n",send_pei_fan.sj2[i].name,
                                send_pei_fan.sj2[i].wna, send_pei_fan.sj2[i].weight);
*/
                        if (send_run_table.mathine==send_pei_fan.sj2[i].mathine)
                            i++;
                          }//end while
                           send_pei_fan.sj2_sum=i;
					}//end if sqlexec
//                    SQLTransact(henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);


 // search  for  yz1
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
                    int i=0;
                    lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.dou,a.fast_do,a.drop_do,a.control_time,a.add_time,a.stop_time,\
a.gon_cha,a.weight,b.mathine from  you_liao_table a,cai_liao_table b where  add_time=1 and  a.cai_number=b.cai_number and b.wna=2 and a.mathine=b.mathine and a.pei_number like \'");
                            lstrcat(select_str,pei_number);
                            lstrcat(select_str,"\' order by a.weight desc");

                    retcode=lstrlen(select_str);
                    //printf("\nlen=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                         while (TRUE)
                         {
							 DWORD  temp;
                        SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.yz1[i].cai_number,6,&temp);
                        SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.yz1[i].name,15,&temp);
                        SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.yz1[i].wna,5,&temp);
                        SQLBindCol(hstmt,4,SQL_C_CHAR,send_pei_fan.yz1[i].dou,6,&temp);
                        SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.yz1[i].fast_do,0,&temp);
                        SQLBindCol(hstmt,6,SQL_C_FLOAT,&send_pei_fan.yz1[i].drop_do,0,&temp);
                        SQLBindCol(hstmt,7,SQL_C_FLOAT,&send_pei_fan.yz1[i].control_time,0,&temp);
                        SQLBindCol(hstmt,8,SQL_C_SLONG,&send_pei_fan.yz1[i].add_time,0,&temp);
                        SQLBindCol(hstmt,9,SQL_C_SLONG,&send_pei_fan.yz1[i].stop_time,0,&temp);

                        SQLBindCol(hstmt,10,SQL_C_FLOAT,&send_pei_fan.yz1[i].gon_cha,0,&temp);
                        SQLBindCol(hstmt,11,SQL_C_FLOAT,&send_pei_fan.yz1[i].weight,0,&temp);
						SQLBindCol(hstmt,12,SQL_C_SLONG,&send_pei_fan.yz1[i].mathine,0,&temp);

                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*                          //printf("%-5s %5s %5s\n",send_pei_fan.yz1[i].name,
                                send_pei_fan.yz1[i].wna, send_pei_fan.yz1[i].dou);
*/
                        if (send_run_table.mathine==send_pei_fan.yz1[i].mathine)
                            i++;
                          }
                           send_pei_fan.yz1_sum=i;
					}//end if sqlexec
//                    SQLTransact(henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);



 // search  for  yz2
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
                    int i=0;
                    lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.dou,a.fast_do,a.drop_do,a.control_time,a.add_time,a.stop_time,\
a.gon_cha,a.weight,b.mathine from  you_liao_table a,cai_liao_table b where  add_time=2 and  a.cai_number=b.cai_number and b.wna=2 and a.mathine=b.mathine and a.pei_number like \'");
                            lstrcat(select_str,pei_number);
                            lstrcat(select_str,"\' order by a.weight desc");

                    retcode=lstrlen(select_str);
                    //printf("\nlen=%d,%s\n",retcode,select_str);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                         while (TRUE)
                         {

                        SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.yz2[i].cai_number,6,NULL);
                        SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.yz2[i].name,15,NULL);
                        SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.yz2[i].wna,5,NULL);
                        SQLBindCol(hstmt,4,SQL_C_CHAR,send_pei_fan.yz2[i].dou,6,NULL);
                        SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.yz2[i].fast_do,0,NULL);
                        SQLBindCol(hstmt,6,SQL_C_FLOAT,&send_pei_fan.yz2[i].drop_do,0,NULL);
                        SQLBindCol(hstmt,7,SQL_C_FLOAT,&send_pei_fan.yz2[i].control_time,0,NULL);
                        SQLBindCol(hstmt,8,SQL_C_SLONG,&send_pei_fan.yz2[i].add_time,0,NULL);
                        SQLBindCol(hstmt,9,SQL_C_SLONG,&send_pei_fan.yz2[i].stop_time,0,NULL);

                        SQLBindCol(hstmt,10,SQL_C_FLOAT,&send_pei_fan.yz2[i].gon_cha,0,NULL);
                        SQLBindCol(hstmt,11,SQL_C_FLOAT,&send_pei_fan.yz2[i].weight,0,NULL);
                        SQLBindCol(hstmt,12,SQL_C_SLONG,&send_pei_fan.yz2[i].mathine,0,NULL);
                      

                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*                          //printf("%-5s %5s %5s\n",send_pei_fan.yz2[i].name,
                                send_pei_fan.yz2[i].wna, send_pei_fan.yz2[i].dou);
*/
                        if (send_run_table.mathine==send_pei_fan.yz2[i].mathine)
                            i++;
                          }
                           send_pei_fan.yz2_sum=i;
					}//end if sqlexec
//                    SQLTransact(henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);



 // search  for  yt1
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
                    int i=0;
                    lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.dou,a.fast_do,a.drop_do,a.control_time,a.add_time,a.stop_time,\
a.gon_cha,a.weight,b.mathine from  you_liao_table a,cai_liao_table b where  add_time=1 and  a.cai_number=b.cai_number and b.wna=3 and a.mathine=b.mathine and a.pei_number like \'");
                            lstrcat(select_str,pei_number);
                            lstrcat(select_str,"\' order by a.weight desc");

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SDWORD  temp;
                         while (TRUE)
                         {

                        SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.yt1[i].cai_number,5,&temp);
                        SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.yt1[i].name,14,&temp);
                        SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.yt1[i].wna,4,&temp);
                        SQLBindCol(hstmt,4,SQL_C_CHAR,send_pei_fan.yt1[i].dou,5,&temp);
                        SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.yt1[i].fast_do,0,NULL);
                        SQLBindCol(hstmt,6,SQL_C_FLOAT,&send_pei_fan.yt1[i].drop_do,0,NULL);
                        SQLBindCol(hstmt,7,SQL_C_FLOAT,&send_pei_fan.yt1[i].control_time,0,NULL);
                        SQLBindCol(hstmt,8,SQL_C_SLONG,&send_pei_fan.yt1[i].add_time,0,NULL);
                        SQLBindCol(hstmt,9,SQL_C_SLONG,&send_pei_fan.yt1[i].stop_time,0,NULL);

                        SQLBindCol(hstmt,10,SQL_C_FLOAT,&send_pei_fan.yt1[i].gon_cha,0,NULL);
                        SQLBindCol(hstmt,11,SQL_C_FLOAT,&send_pei_fan.yt1[i].weight,0,NULL);
						SQLBindCol(hstmt,12,SQL_C_SLONG,&send_pei_fan.yt1[i].mathine,0,NULL);

                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*                          //printf("%-5s %5s %5s\n",send_pei_fan.yt1[i].name,
                                send_pei_fan.yt1[i].wna, send_pei_fan.yt1[i].dou);
*/
                        if (send_run_table.mathine==send_pei_fan.yt1[i].mathine)
                                        i++;
                          }
                           send_pei_fan.yt1_sum=i;
					}//end if sqlexec
//                    SQLTransact(henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);



 // search  for  yt2
             retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
             if (retcode == SQL_SUCCESS)
                {
                    int i=0;
                    lstrcpy(select_str,"select  a.cai_number,\
b.name,b.wna,b.dou,a.fast_do,a.drop_do,a.control_time,a.add_time,a.stop_time,\
a.gon_cha,a.weight,b.mathine from  you_liao_table a,cai_liao_table b where  add_time=2 and  a.cai_number=b.cai_number and b.wna=3 and a.mathine=b.mathine and a.pei_number like \'");
                            lstrcat(select_str,pei_number);
                            lstrcat(select_str,"\' order by a.weight desc");


                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SDWORD  temp;
                         while (TRUE)
                         {

                        SQLBindCol(hstmt,1,SQL_C_CHAR,send_pei_fan.yt2[i].cai_number,5,&temp);
                        SQLBindCol(hstmt,2,SQL_C_CHAR,send_pei_fan.yt2[i].name,14,&temp);
                        SQLBindCol(hstmt,3,SQL_C_CHAR,send_pei_fan.yt2[i].wna,4,&temp);
                        SQLBindCol(hstmt,4,SQL_C_CHAR,send_pei_fan.yt2[i].dou,5,&temp);
                        SQLBindCol(hstmt,5,SQL_C_FLOAT,&send_pei_fan.yt2[i].fast_do,0,NULL);
                        SQLBindCol(hstmt,6,SQL_C_FLOAT,&send_pei_fan.yt2[i].drop_do,0,NULL);
                        SQLBindCol(hstmt,7,SQL_C_FLOAT,&send_pei_fan.yt2[i].control_time,0,NULL);
                        SQLBindCol(hstmt,8,SQL_C_SLONG,&send_pei_fan.yt2[i].add_time,0,NULL);
                        SQLBindCol(hstmt,9,SQL_C_SLONG,&send_pei_fan.yt2[i].stop_time,0,NULL);

                        SQLBindCol(hstmt,10,SQL_C_FLOAT,&send_pei_fan.yt2[i].gon_cha,0,NULL);
                        SQLBindCol(hstmt,11,SQL_C_FLOAT,&send_pei_fan.yt2[i].weight,0,NULL);
						SQLBindCol(hstmt,12,SQL_C_SLONG,&send_pei_fan.yt2[i].mathine,0,NULL);

                            retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    break;      //exit  while
/*                          //printf("%-5s %5s %5s\n",send_pei_fan.yt2[i].name,
                                send_pei_fan.yt2[i].wna, send_pei_fan.yt2[i].dou);
*/
                        if (send_run_table.mathine==send_pei_fan.yt2[i].mathine)                            
                                        i++;

                          }
                           send_pei_fan.yt2_sum=i;
					}//end if sqlexec
//                    SQLTransact(henv, ljxt_hdbc, SQL_COMMIT);
					SQLTransact(ljxt_henv, ljxt_hdbc, SQL_COMMIT);
                   }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

	control_can_read =CAN_READ;   //可读
    can_run_flag=(flag>>4);  //CAN_RUN;   //可读
	if (can_run_flag!=3)
		Prev_Run_Flag=can_run_flag;
	return 1;
}

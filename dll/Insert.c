#include <time.h>
#include <stdio.h>
#include <sys/types.h>
#include <sys/timeb.h>
  #include "windows.h"
  #include "SQL.H"
  #include "SQLext.H"
  #include <string.h> 
  #include "dll.h"
  #include "temp.h"
             
	#pragma     data_seg("sdata")
	  int  Pai_Fan_Run_Flag=0;
	  int  make_data_open_flag=0;
	  HENV make_henv=NULL;
	  HDBC make_hdbc=NULL; 
	  struct make_lian_liao  get_make_lian_liao[20]={0,0,0,0};
	  struct  Fen_Disp_Dou  Chans_Data={0,0,0};
	#pragma data_seg()
	  extern float Round(float,int);
//每段写一次
//全、局变量    duan_begin_time   ---->c1.c   write_power


int  WINAPI   write_make(struct make_lian_liao temp)
{
        int   _duan;    
		char  power_buffer[20];
		char tempro_buffer[14];
		int i;
        _duan=temp.duan_hao-1;
		if (_duan>20 )
		{
			MessageBox(0,"错误","段号超过20 只取20",MB_OK);
			_duan=19;
		}
        get_make_lian_liao[_duan].duan_hao=_duan+1;
        get_make_lian_liao[_duan].set_hun_time=send_pei_fan.huan[_duan].lian_time;
        get_make_lian_liao[_duan].set_you_time=send_pei_fan.huan[_duan].dou_you_time;
        get_make_lian_liao[_duan].set_tan_time=send_pei_fan.huan[_duan].dou_tan_time;
        get_make_lian_liao[_duan].set_jiao_time=send_pei_fan.huan[_duan].dou_jiao_time;
        get_make_lian_liao[_duan].set_xiao_time=send_pei_fan.huan[_duan].dou_xiao_time;
		get_make_lian_liao[_duan].set_speed=send_pei_fan.huan[_duan].speed;

        get_make_lian_liao[_duan].get_hun_time=temp.get_hun_time;//段时间
        get_make_lian_liao[_duan].get_you_time=temp.get_you_time;
        get_make_lian_liao[_duan].get_tan_time=temp.get_tan_time;
        get_make_lian_liao[_duan].get_jiao_time=temp.get_jiao_time;
        get_make_lian_liao[_duan].get_xiao_time=temp.get_xiao_time;
        get_make_lian_liao[_duan].next_tempro=temp.next_tempro;
		get_make_lian_liao[_duan].next_power=temp.next_power;
        lstrcpy(get_make_lian_liao[_duan].duan_begin_time,temp.duan_begin_time);
		lstrcpy(get_make_lian_liao[_duan].duan_end_time,temp.duan_end_time);

		 
		sprintf(power_buffer,"%6.2f",work_power.power[0]);
		sprintf(tempro_buffer,"%3d",work_power.tempro[0]);
		lstrcpy(get_make_lian_liao[_duan].power,power_buffer);
		lstrcpy(get_make_lian_liao[_duan].tempro,tempro_buffer);

		for(i=1;i<work_power.ptr;i++)
		{
				sprintf(power_buffer,",%6.2f",work_power.power[i]);
				sprintf(tempro_buffer,",%3d",work_power.tempro[i]);
				lstrcat(get_make_lian_liao[_duan].power,power_buffer);
				lstrcat(get_make_lian_liao[_duan].tempro,tempro_buffer);
		}

         work_power.ptr=0;
		 work_power.flag=-1;

		if(temp.get_you_time!=0)
		{
			lstrcpy(get_make_lian_liao[_duan].Chian_Str,"投油1");
			lstrcpy(get_make_lian_liao[_duan].Eng_Str,"dou you1");
			get_make_lian_liao[_duan].Get_the_time=temp.get_you_time;			
		}
		if(temp.get_you2_time!=0)
		{
			lstrcpy(get_make_lian_liao[_duan].Chian_Str,"投油2");
			lstrcpy(get_make_lian_liao[_duan].Eng_Str,"dou you2");
			get_make_lian_liao[_duan].Get_the_time=temp.get_you2_time;			
		}

		if(temp.get_jiao_time!=0)
		{
			lstrcpy(get_make_lian_liao[_duan].Chian_Str,"投胶");
			lstrcpy(get_make_lian_liao[_duan].Eng_Str,"dou jiao");
			get_make_lian_liao[_duan].Get_the_time=temp.get_jiao_time;			
		}
		if(temp.get_tan_time!=0)
		{
			lstrcpy(get_make_lian_liao[_duan].Chian_Str,"投炭");
			lstrcpy(get_make_lian_liao[_duan].Eng_Str,"dou tan");
			get_make_lian_liao[_duan].Get_the_time=temp.get_tan_time;			
		}
		if(temp.get_xiao_time!=0)
		{
			lstrcpy(get_make_lian_liao[_duan].Chian_Str,"投小料");
			lstrcpy(get_make_lian_liao[_duan].Eng_Str,"dou small");
			get_make_lian_liao[_duan].Get_the_time=temp.get_xiao_time;			
		}

	return TRUE;
}
/*
	第一次写库
*/
int First_Dll_Write_DB(struct run_table run_do)
{
    char  select_str[400];
	char buffer[10];
	HSTMT hstmt; 
	RETCODE retcode;
	int i=0;

	if(make_mdb_open_flag==0)
	{
		make_data_open_flag=1;
		retcode = SQLAllocEnv(&make_henv); /* Environment handle */ 
		if (retcode == SQL_SUCCESS) 
		{ 
			retcode = SQLAllocConnect(make_henv, &make_hdbc); /* Connection handle */
			if (retcode == SQL_SUCCESS)
			{
            SQLSetConnectOption(make_hdbc, SQL_LOGIN_TIMEOUT, 5);
            retcode = SQLConnect(make_hdbc, (unsigned char *) "make", SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
			if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
			{
				write_error(ODBC_ERROR,"没设置make ODBC");				
				SQLFreeEnv(make_henv);
				make_mdb_open_flag=0;
				return FALSE;
			}// end if SQLConnect
		}//end if allocconnect
	}//end if allocenv
}//end if make_data_open_flag

                
                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					
					Static_make_index.Sort_Idx=0;
/*
	找相同记录
*/
					lstrcpy(select_str," select Sort_Idx  from make_index \
where  mathine=");
					_itoa(run_do.mathine,buffer, 10 );
					lstrcat(select_str,buffer);
					lstrcat(select_str," and m_number like '");
					lstrcat(select_str,run_do._number);
					lstrcat(select_str,"' and name like '");
					lstrcat(select_str,run_do.name);
					Get_Date((Static_make_index.begin_date));					
					lstrcat(select_str,"' and begin_date=#");
					lstrcat(select_str,Static_make_index.begin_date);
					lstrcat(select_str,"# and ban=");
					_itoa(run_do.ban,buffer, 10 );
					lstrcat(select_str,buffer);										

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)//
                    {
                        
						DWORD   temp_int;						
						SQLBindCol(hstmt, 1, SQL_C_SLONG,&(Static_make_index.Sort_Idx), 0, &temp_int);
                        retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
									Static_make_index.Sort_Idx=0;
                     }//exec
				}
                SQLFreeStmt(hstmt, SQL_DROP);
		
/*		找最大索引
*/

			if(Static_make_index.Sort_Idx==0)
			{
                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {


					lstrcpy(select_str," select max(Sort_Idx) from make_index \
where  mathine=");
					_itoa(run_do.mathine,buffer, 10 );
					lstrcat(select_str,buffer);
                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        
						DWORD   temp_int;						
						SQLBindCol(hstmt, 1, SQL_C_SLONG,&(Static_make_index.Sort_Idx), 0, &temp_int);
                        retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    Static_make_index.Sort_Idx=1;
							else
									Static_make_index.Sort_Idx++;						
                     }//exec
				}
                SQLFreeStmt(hstmt, SQL_DROP);						
			}
			else
				return TRUE;
				//end if sort_idx

				lstrcpy(Static_make_index.m_number,run_do._number);
				lstrcpy(Static_make_index.name,run_do.name);
				Static_make_index.ban=run_do.ban;
				Static_make_index.mathine=run_do.mathine;
				Static_make_index.che=run_do.total_che;

				Get_Date((Static_make_index.begin_date));					
				Get_Time((Static_make_index.begin_time));					


                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {								
		   lstrcpy(select_str," insert into make_index \
(Sort_Idx,mathine,m_number,name,ban,che,flag,begin_date,begin_time\
)  values(?,?,?,?,?,?,1,'");//  01/01/11','12:12:10')"); //total 12 ?  duan_hao
	   lstrcat(select_str,Static_make_index.begin_date);
		   lstrcat(select_str,"','");
		   lstrcat(select_str,Static_make_index.begin_time);
		   lstrcat(select_str,"')");

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
						SQLINTEGER      cbSort_Idx = 0, cbMathine= 0, cbChe = 0, cbBegin_Time = SQL_NTS,
										cbBegin_Date = SQL_NTS,cbName = SQL_NTS,cbNumber= SQL_NTS;

                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &(Static_make_index.Sort_Idx), 0, &cbSort_Idx);
						SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_index.mathine), 0, &cbMathine);
                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_VARCHAR, 6, 0, Static_make_index.m_number, 0,&cbNumber);
						SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_VARCHAR, 15, 0, Static_make_index.name, 0, &cbName);
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &(Static_make_index.ban), 0, NULL);
						SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &(Static_make_index.che), 0, NULL);//che						
                        retcode = SQLExecute(hstmt);         /* Execute statement with */

						
                    }//end insert

                }//alloc stmt
                SQLFreeStmt(hstmt, SQL_DROP);
if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
		return FALSE;
	return TRUE;

}


int   First_Write_Che(struct run_table run_do)
{
    char  select_str[400];
	HSTMT hstmt; 
	RETCODE retcode;
    int i=0;

	if(make_data_open_flag==0)
	{
				write_error(ODBC_ERROR,"没设置make ODBC");				
				return FALSE;
	}

//found  if  sort_idx have runing
				Static_make_table.only_key=0;
                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					char buffer[20];
					lstrcpy(select_str," select only_key,now_che from make_table \
where  mathine=");
					_itoa((int)(run_do.mathine),buffer, 10 );
					lstrcat(select_str,buffer);
					lstrcat(select_str," and sort_idx=");
					_itoa((int)(Static_make_index.Sort_Idx),buffer, 10 );
					lstrcat(select_str,buffer);
					lstrcat(select_str," order by now_che desc");
                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        
						DWORD   temp_int;
						SQLBindCol(hstmt, 1, SQL_C_SLONG, &(Static_make_table.only_key), 0, &temp_int);
						SQLBindCol(hstmt, 2, SQL_C_SLONG, &(Static_make_table.now_che), 0, &temp_int);
                        retcode = SQLFetch(hstmt);
//                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
//                                    Static_make_table.only_key=1;
//							else
//									Static_make_table.only_key++;						
                     }//exec
                }//alloc stmt
                SQLFreeStmt(hstmt, SQL_DROP);

	if(Static_make_table.only_key==0)
	{
                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					char buffer[20];
					lstrcpy(select_str," select max(only_key) from make_table \
where  mathine=");
					_itoa((int)(run_do.mathine),buffer, 10 );
					lstrcat(select_str,buffer);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        
						DWORD   temp_int;
						SQLBindCol(hstmt, 1, SQL_C_SLONG, &(Static_make_table.only_key), 0, &temp_int);
                        retcode = SQLFetch(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
                                    Static_make_table.only_key=1;
							else
									Static_make_table.only_key++;						
                     }//exec
                }//alloc stmt
                SQLFreeStmt(hstmt, SQL_DROP);
				Static_make_table.now_che=1;	//run_do.now_che;
	}
	else//update  now_che
	{
				Static_make_table.now_che++;
	}

//next doing  Insert make_table
				Static_make_table.mathine=run_do.mathine;
				Static_make_table.Sort_Idx=Static_make_index.Sort_Idx;

				Get_Date(Static_make_table.begin_date);					
				Get_Time(Static_make_table.begin_time);					

                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					char buffer[20];

		   lstrcpy(select_str," insert into  make_table \
(Sort_Idx,mathine,only_key,now_che,m_date,m_time\
) values(?,?,?,?,'");//?,?) "); //total 12 ?  duan_hao
	   lstrcat(select_str,Static_make_table.begin_date);
		   lstrcat(select_str,"','");
		   lstrcat(select_str,Static_make_table.begin_time);
		   lstrcat(select_str,"')");

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &(Static_make_table.Sort_Idx), 0, NULL);
						SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.mathine), 0, NULL);
						SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &(Static_make_table.only_key), 0, NULL);
						SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &(Static_make_table.now_che), 0, NULL);

                            retcode = SQLExecute(hstmt);         /* Execute statement with */

                    }//end insert

                }//alloc stmt
                SQLFreeStmt(hstmt, SQL_DROP);
  if (retcode == SQL_SUCCESS)
	return TRUE;
  return FALSE;
		                  

}

int WINAPI   write_che(struct make_base_type temp)
{
    char  select_str[400];
	HSTMT hstmt; 
	RETCODE retcode;
    int i=0;
    char  tmpbuf[128];
	int           temp_wna,temp_dou;
	SQLINTEGER    temp_int=0,cbchar= SQL_NTS;

	if (can_run_flag==0||can_run_flag==2)			//非实际运行
	{
			write_error(ODBC_ERROR,"非运行态不能写盘");				
			return FALSE;	
	}
	if(Pai_Fan_Run_Flag==0)
	{
		if(First_Dll_Write_DB(send_run_table)==FALSE) return FALSE;	//第一次写盘  返回 Sort_Idx,Only_key
		Pai_Fan_Run_Flag=1;
	}
	if (First_Write_Che(send_run_table)<0)
		return FALSE;

                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
//						DWORD temp_int;

                        for(i=0;i<temp.th1_sum;i++)
                        {
	                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &(Static_make_table.only_key), 0, &temp_int);
						    SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.th1[i].mathine), 0, &temp_int);

	
							SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.th1[i].add_time,0,&temp_int);  //give_time
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.th1[i].cai_number,0,&cbchar);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.th1[i].name,0,&cbchar);
							temp_wna=(*temp.th1[i].wna-'0');
							temp_dou=(*temp.th1[i].dou-'0');
														
							
							temp.th1[i].get_weight=Round(temp.th1[i].get_weight,2);
							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);
							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,8, 0, &(temp.th1[i].weight),0,&temp_int);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,8, 2, &(temp.th1[i].get_weight),0,&temp_int);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.th1[i].gon_cha),0,&temp_int);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
	                        if (retcode!=SQL_SUCCESS) break;
						}//end for i
                      for(i=0;i<temp.th2_sum;i++)
                        {
	                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.only_key), 0, NULL);
		                    SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.th2[i].mathine, 0, NULL);

							SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.th2[i].add_time,0,NULL);  //give_time
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.th2[i].cai_number,0,NULL);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.th2[i].name,0,NULL);
							temp_wna=(*temp.th2[i].wna-'0');
							temp_dou=(*temp.th2[i].dou-'0');

							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);
							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,8, 2, &(temp.th2[i].weight),0,NULL);
							
							temp.th2[i].get_weight=Round(temp.th2[i].get_weight,2);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,8, 2, &(temp.th1[i].get_weight),0,NULL);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.th2[i].gon_cha),0,NULL);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }//end for i
                        for(i=0;i<temp.sj1_sum;i++)
                        {
							SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.only_key), 0, NULL);							
							SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.sj1[i].mathine, 0, NULL);
	
	                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.sj1[i].add_time,0,NULL);  //give_time
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.sj1[i].cai_number,0,&cbchar);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.sj1[i].name,0,&cbchar);
							temp_wna=4;//(*temp.sj1[i].wna-'0');
							temp_dou=0;//(*temp.sj1[i].dou-'0');

							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);

							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.sj1[i].weight),0,&temp_int);
							temp.sj1[i].get_weight=Round(temp.sj1[i].get_weight,2);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,8, 2, &(temp.sj1[i].get_weight),0,&temp_int);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.sj1[i].gon_cha),0,&temp_int);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                        for(i=0;i<temp.sj2_sum;i++)
                        {
							SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.only_key), 0, NULL);
							SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.sj2[i].mathine, 0, NULL);
	
	                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.sj2[i].add_time,0,NULL);  //give_time
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.sj2[i].cai_number,0,&cbchar);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.sj2[i].name,0,&cbchar);
							temp_wna=4;//(*temp.sj1[i].wna-'0');
							temp_dou=0;//(*temp.sj1[i].dou-'0');

							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);

							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.sj2[i].weight),0,&temp_int);
							temp.sj2[i].get_weight=Round(temp.sj2[i].get_weight,2);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.sj2[i].get_weight),0,&temp_int);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.sj2[i].gon_cha),0,&temp_int);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
	                        if (retcode!=SQL_SUCCESS) break;
						}
                        for(i=0;i<temp.yz1_sum;i++)
                        {
	                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.only_key), 0, NULL);
	                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yz1[i].mathine, 0, NULL);
				//         SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &Static_make_table.duan_hao, 0,NULL );
	
							SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yz1[i].add_time,0,NULL);
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.yz1[i].cai_number,0,NULL);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yz1[i].name,0,NULL);
							temp_wna=(*temp.yz1[i].wna-'0');
							temp_dou=(*temp.yz1[i].dou-'0');

							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);
							temp.yz1[i].get_weight=Round(temp.yz1[i].get_weight,2);
							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yz1[i].weight),0,NULL);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yz1[i].get_weight),0,NULL);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yz1[i].gon_cha),0,NULL);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                        for(i=0;i<temp.yz2_sum;i++)
                        {
	                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.only_key), 0, NULL);
							SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yz2[i].mathine, 0, NULL);
	//                      SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &Static_make_table.now_che, 0,NULL );
	
							SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yz2[i].add_time,0,NULL);
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.yz2[i].cai_number,0,NULL);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yz2[i].name,0,NULL);
							temp_wna=(*temp.yz2[i].wna-'0');
							temp_dou=(*temp.yz2[i].dou-'0');
							temp.yz2[i].get_weight=Round(temp.yz2[i].get_weight,2);
							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);

							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yz2[i].weight),0,NULL);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yz2[i].get_weight),0,NULL);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT,SQL_REAL,0, 0, &(temp.yz2[i].gon_cha),0,NULL);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                        for(i=0;i<temp.yt1_sum;i++)
                        {
	                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.only_key), 0, NULL);
							SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.yt1[i].mathine), 0, NULL);
	//                      SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &Static_make_table.now_che, 0,NULL );

							SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yt1[i].add_time,0,NULL);
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.yt1[i].cai_number,0,NULL);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yt1[i].name,0,NULL);
							temp_wna=(*temp.yt1[i].wna-'0');
							temp_dou=(*temp.yt1[i].dou-'0');

							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);
							temp.yt1[i].get_weight=Round(temp.yt1[i].get_weight,2);

							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yt1[i].weight),0,NULL);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yt1[i].get_weight),0,NULL);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yt1[i].gon_cha),0,NULL);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }
                        for(i=0;i<temp.yt2_sum;i++)
                        {
	                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.only_key), 0, NULL);
							SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(temp.yt1[i].mathine), 0, NULL);
//                      SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &Static_make_table.now_che, 0,NULL );

							SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.yt2[i].add_time,0,NULL);
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,6, 0, temp.yt2[i].cai_number,0,NULL);
							SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR,15, 0, temp.yt2[i].name,0,NULL);
							temp_wna=(*temp.yt2[i].wna-'0');
							temp_dou=(*temp.yt2[i].dou-'0');

							SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_wna,0,&temp_int);
							SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT,0, 0, &temp_dou,0,&temp_int);

							SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yt2[i].weight),0,NULL);
							temp.yt2[i].get_weight=Round(temp.yt2[i].get_weight,2);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yt2[i].get_weight),0,NULL);
							SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL,0, 0, &(temp.yt2[i].gon_cha),0,NULL);
							SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che), 0, &temp_int);
                            retcode = SQLExecute(hstmt);         /* Execute statement with */
                            if (retcode!=SQL_SUCCESS) break;
                        }

						
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

/*                retcode = SQLAllocStmt(make_hdbc, &hstmt); 
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 11?  duan_hao
                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);


// insert into sj1
                retcode = SQLAllocStmt(make_hdbc, &hstmt); 
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao



                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);
// insert into sj2
/*				
                retcode = SQLAllocStmt(make_hdbc, &hstmt); 
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao


                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        }//end for
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);
*/
// insert into yz1
/*                retcode = SQLAllocStmt(make_hdbc, &hstmt); 
                if (retcode == SQL_SUCCESS)
                {

   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao


                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {

                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);
// insert into yz2
/*                retcode = SQLAllocStmt(make_hdbc, &hstmt); 
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao


                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);


// insert into yt1
                retcode = SQLAllocStmt(make_hdbc, &hstmt); 
                if (retcode == SQL_SUCCESS)
                {

   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao


                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {

                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);

			// insert into yt2
/*                retcode = SQLAllocStmt(make_hdbc, &hstmt); 
                if (retcode == SQL_SUCCESS)
                {

   lstrcpy(select_str," insert into make_pei_fan \
(only_key,mathine,give_time,cai_number,cai_name,\
wna,dou,w1,w2,set_gon_cha,now_che)  values(?,?,?,?,?,?,?,?,?,?,?) "); //total 12 ?  duan_hao


                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);
*/

				//insert into make_lian_liao
                retcode = SQLAllocStmt(make_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert into make_base\
(only_key,mathine,set_weight,get_weight,set_pai_time,get_pai_time,\
pai_tempro,huan_time,all_power,now_che)  values(?,?,?,?,?,?,?,?,?,?)"); //total 9 ?  duan_hao


                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_INTEGER, 0, 0, &(Static_make_table.only_key), 0, NULL);
                        SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &Static_make_table.mathine, 0,NULL );

                        SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_REAL, 0, 0, &send_pei_fan.base.weight, 0,NULL );
                        SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_REAL, 0, 0, &temp.get_weight, 0,NULL );
                        SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &send_pei_fan.base.pai_liao_time, 0,NULL );
                        SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.get_pai_time, 0,NULL );                        
                        SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_REAL, 0, 0, &temp.pai_tempro, 0,NULL );
						SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &temp.huan_time, 0,NULL );//总混炼时间
						SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL, 0, 0, &temp.all_power, 0,NULL );
						SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che),0,NULL);
                        retcode = SQLExecute(hstmt);         /* Execute statement with */
                        
                    }//end insert
                }// END IF SQLAllocStmt
            SQLFreeStmt(hstmt, SQL_DROP);


/*
			写段数据
			图形为一车

*/
                retcode = SQLAllocStmt(make_hdbc, &hstmt);
                if (retcode == SQL_SUCCESS)
                {
   lstrcpy(select_str," insert into make_lian_liao \
(only_key,now_che,duan_hao, mathine,\
China_Str,Eng_Str,\
power,tempro,duan_begin_time,duan_end_time,next_tempro,next_power,set_speed,get_hun_time) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); //total 12 ?
   

                    retcode = SQLPrepare(hstmt, select_str, SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
						int i=0;
						for(i=0;i<temp.pei_max;i++)
						{
                            SQLBindParameter(hstmt, 1, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_INTEGER, 0, 0, &(Static_make_table.only_key), 0, &temp_int);
                            SQLBindParameter(hstmt, 2, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.now_che),0,&temp_int);
                            SQLBindParameter(hstmt, 3, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_lian_liao[i].duan_hao),0,&temp_int);
							SQLBindParameter(hstmt, 4, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(Static_make_table.mathine),0,&temp_int);

                            SQLBindParameter(hstmt, 5, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 18, 0, get_make_lian_liao[i].Chian_Str,0,&cbchar);
                            SQLBindParameter(hstmt, 6, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 18, 0, get_make_lian_liao[i].Eng_Str,0,&cbchar);
                            SQLBindParameter(hstmt, 7, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_LONGVARCHAR, 2000, 0, get_make_lian_liao[i].power,0,&cbchar);
                            SQLBindParameter(hstmt, 8, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_LONGVARCHAR, 2000, 0, get_make_lian_liao[i].tempro,0,&cbchar);
							SQLBindParameter(hstmt, 9, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_TIME, 10, 0, get_make_lian_liao[i].duan_begin_time,0,NULL);
                            SQLBindParameter(hstmt, 10, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_TIME, 10, 0, get_make_lian_liao[i].duan_end_time,0,NULL);

                            SQLBindParameter(hstmt, 11, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_lian_liao[i].next_tempro),0,&temp_int);
							SQLBindParameter(hstmt, 12, SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL, 0, 0, &(get_make_lian_liao[i].next_power),0,&temp_int);
							SQLBindParameter(hstmt, 13, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_lian_liao[i].set_speed),0,&temp_int);
							SQLBindParameter(hstmt, 14, SQL_PARAM_INPUT, SQL_C_SSHORT, SQL_SMALLINT, 0, 0, &(get_make_lian_liao[i].get_hun_time),0,&temp_int);//段时间
                            retcode = SQLExecute(hstmt);
                            if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)  break;
                        }
                    }//end insert
                }// END IF SQLAllocStmt				
				SQLFreeStmt(hstmt, SQL_DROP);

   if (retcode == SQL_SUCCESS || retcode == SQL_SUCCESS_WITH_INFO)	
			return TRUE;
	return FALSE;
}


int WINAPI write_chans(int mathine,int dou, int kind,int send_data)
{
		Chans_Data.mathine=mathine;
		Chans_Data.dou=dou;
		Chans_Data.kind=kind;
		Chans_Data.send_data=send_data;
		return 1;
}
int WINAPI write_runfch(int now_data,float one, float two,float power,int time, char * statu)
{
		Chans_Data.now_data=now_data;
		Chans_Data.one=one;
		Chans_Data.two=two;
		Chans_Data.power=power;
		Chans_Data.my_time=time;
		lstrcpy(Chans_Data.statu,statu);
		return 1;
}
int WINAPI read_fen_statu(struct  Fen_Disp_Dou * x)
{
		x->mathine=Chans_Data.mathine;
		x->dou=Chans_Data.dou;
		x->kind=Chans_Data.kind;
		x->send_data=Chans_Data.send_data;

		x->now_data=Chans_Data.now_data;
		x->one=Chans_Data.one;
		x->two=Chans_Data.two;
		x->power=Chans_Data.power;
		x->my_time=Chans_Data.my_time;
		lstrcpy(x->statu,Chans_Data.statu);
		return 1;
}

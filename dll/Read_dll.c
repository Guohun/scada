#include <windows.h>
#include "SQL.H"
#include "SQLext.H"

#include "dll.h"
#include "temp.h"
#pragma     data_seg("sdata")
		int   comm_mdb_open_flag=0; /* 0---没打开* 1--打开*/
		HENV comm_henv=NULL;
		HDBC comm_hdbc=NULL; 	
		int  Prev_Run_Flag=0; 	
		char grap_name[20]="";
#pragma data_seg()


//dll->控制程序
 int WINAPI read_produce(struct Produce_type  *temp)
{
	char *pdest;
    int i;
    
    if (control_can_read!=CAN_READ) return WRITE_DOING;
	if (can_run_flag==STOP_DOING)
			
		return  STOP_DOING;
	
		


   /* send_run_table--->temp*/
    temp->sj1_sum=send_pei_fan.sj1_sum;
    temp->sj2_sum=send_pei_fan.sj2_sum;

    temp->th1_sum=send_pei_fan.th1_sum;
    temp->th2_sum=send_pei_fan.th2_sum;

    temp->yz1_sum=send_pei_fan.yz1_sum;
    temp->yz2_sum=send_pei_fan.yz2_sum;

    temp->yt1_sum=send_pei_fan.yt1_sum;
    temp->yt2_sum=send_pei_fan.yt2_sum;

    temp->pei_max=send_pei_fan.pei_max;
    memcpy((void *)&(temp->base),(void *)&(send_pei_fan.base),sizeof(send_pei_fan.base) );
	for(i=0;i<send_pei_fan.th1_sum;i++)
		memcpy((void *)&(temp->th1[i]), (void *)&(send_pei_fan.th1[i]),sizeof(send_pei_fan.th1[i]));
	for(i=0;i<send_pei_fan.th2_sum;i++)
		memcpy((void *)&(temp->th2[i]), (void *)&send_pei_fan.th2[i],sizeof(send_pei_fan.th2[i]));
	for(i=0;i<send_pei_fan.sj1_sum;i++)
		memcpy((void *)&(temp->sj1[i]), (void *)&send_pei_fan.sj1[i],sizeof(send_pei_fan.sj1[i]));

    for(i=0;i<send_pei_fan.sj2_sum;i++)
        memcpy((void *)&(temp->sj2[i]), (void *)&(send_pei_fan.sj2[i]),sizeof(send_pei_fan.sj2[i]));


    for(i=0;i<send_pei_fan.yz1_sum;i++)
        memcpy((void *)&(temp->yz1[i]), (void *)&(send_pei_fan.yz1[i]),sizeof(send_pei_fan.yz1[i]));

    for(i=0;i<send_pei_fan.yz2_sum;i++)
        memcpy((void *)&(temp->yz2[i]), (void *)&(send_pei_fan.yz2[i]),sizeof(send_pei_fan.yz2[i]));

    for(i=0;i<send_pei_fan.yt1_sum;i++)
        memcpy((void *)&(temp->yt1[i]), (void *)&(send_pei_fan.yt1[i]),sizeof(send_pei_fan.yt1[i]));

    for(i=0;i<send_pei_fan.yt2_sum;i++)
        memcpy((void *)&(temp->yt2[i]), (void *)&(send_pei_fan.yt2[i]),sizeof(send_pei_fan.yt2[i]));

    for(i=0;i<send_pei_fan.pei_max;i++)
        memcpy((void *)&(temp->huan[i]), (void *)&(send_pei_fan.huan[i]),sizeof(send_pei_fan.huan[i]));

/*
    把send_run_table中数据送入  produce_DO
    返回 1
*/
    return control_can_read;
}

// dll->控制程序
 int WINAPI  read_run(struct run_table* x,int *run_flag)
  {
      if (can_run_flag==STOP_DOING)
	{
		*run_flag=STOP_DOING;
		return  STOP_DOING;
	}
 
	   *run_flag=can_run_flag;
//        memcpy(x, (void *)&send_run_table,sizeof(struct run_table));
		
		lstrcpy(x->name,send_run_table.name);
		lstrcpy(x->_number,send_run_table._number);
		x->total_che=send_run_table.total_che;
		if (can_run_flag!=3)
		{
			x->now_che=send_run_table.now_che;	/*第几车*/
			x->all_duan=send_pei_fan.pei_max;//send_run_table.all_duan;/* 总段 */
			x->ban=send_run_table.ban;		/*班  */
			x->mathine=send_run_table.mathine;  /*机号 */
			x->all_time=send_run_table.all_time;	  /*总工作时间 */
			x->all_che_time=send_run_table.all_che_time;	/*总一车时间 */
			x->duan_time=send_run_table.duan_time; /*段时间 */
		}
		else
		{
			x->now_che=send_run_table.now_che;	/*第几车*/
			x->all_duan=send_pei_fan.pei_max;//send_run_table.all_duan;/* 总段 */
//			x->ban=send_run_table.ban;		/*班  */
//			x->mathine=send_run_table.mathine;  /*机号 */
			x->all_time=send_run_table.all_time;	  /*总工作时间 */
			x->all_che_time=send_run_table.all_che_time;	/*总一车时间 */
			x->duan_time=send_run_table.duan_time; /*段时间 */

		}
		//MessageBox(0,"asdf","adsfa",MB_OK);
		
        
/*  
  if (can_run_flag==STOP) return STOP_DOING;
   *run_flag = CAN_RUN;
   IF(can_read !=CAN_READ) return WRITE_Doing;
   把send_run_table中数据送入  
   返回 1
*/
    return 1;
}
WINMMAPI int WINAPI read_send_table(int code,int mathine,struct send_table *x)
{
    char  select_str[400];
	HSTMT hstmt; 
	RETCODE retcode;

	int flag=0;
if (comm_mdb_open_flag==0)
{
	comm_mdb_open_flag=1;
	retcode = SQLAllocEnv(&comm_henv); /* Environment handle */ 
	if (retcode == SQL_SUCCESS) 
	{ 
        retcode = SQLAllocConnect(comm_henv, &comm_hdbc); /* Connection handle */
        if (retcode == SQL_SUCCESS)
        {
            SQLSetConnectOption(comm_hdbc, SQL_LOGIN_TIMEOUT, 5);
            retcode = SQLConnect(comm_hdbc, (unsigned char *) "comm", SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
			if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
			{
				write_error(ODBC_ERROR,"没设置comm ODBC");				
				SQLFreeEnv(comm_henv);
				comm_henv=NULL;
				comm_mdb_open_flag=0;
				return FALSE;
			}			
		}//end SQLALLOCConnect
	}//end SQLALLocEnv
};//end if comm_mdb_open_flag
/*
               SQLDisconnect(hdbc);
            }// end if SQLConnect
            SQLFreeConnect(hdbc);
        }//end if allocconnect
        SQLFreeEnv(henv);
    }//end if allocenv

	if (retcode==SQL_SUCCESS||retcode==SQL_SUCCESS_WITH_INFO) 					
			return TRUE;	
	else
			return FALSE;
*/	

                retcode = SQLAllocStmt(comm_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					char buffer[20];
					lstrcpy(select_str," select * from  send_table \
where  mathine=");
					_itoa(mathine,buffer, 10 );
					lstrcat(select_str,buffer);
					lstrcat(select_str," and code=");
					_itoa(code,buffer, 10 );
					lstrcat(select_str,buffer);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    if (retcode == SQL_SUCCESS)
                    {
                        
						DWORD   temp_int;						
						SQLBindCol(hstmt, 1, SQL_C_SLONG,&(x->code), 0, &temp_int);
						SQLBindCol(hstmt, 2, SQL_C_CHAR,x->name, 18, &temp_int);
						SQLBindCol(hstmt, 3, SQL_C_FLOAT,&(x->max), 0, &temp_int);
						SQLBindCol(hstmt, 4, SQL_C_FLOAT,&(x->min), 0, &temp_int);
						SQLBindCol(hstmt, 5, SQL_C_SLONG,&(x->AD_max), 0, &temp_int);
						SQLBindCol(hstmt, 6, SQL_C_SLONG,&(x->AD_min), 0, &temp_int);
						SQLBindCol(hstmt, 7, SQL_C_SSHORT,&(x->port), 0, &temp_int);
//						SQLBindCol(hstmt, 7, SQL_C_CHAR,&(x->time), 0, &temp_int);
//						SQLBindCol(hstmt, 8, SQL_C_CHAR,&(x->true_data), 0, &temp_int);
						SQLBindCol(hstmt, 9, SQL_C_SSHORT,&(x->mathine), 0, &temp_int);
//						SQLBindCol(hstmt, 10, SQL_C_SSHORT,&(x->true_AD), 0, &temp_int);
						SQLBindCol(hstmt, 12, SQL_C_SSHORT,&(x->state), 0, &temp_int);
						SQLBindCol(hstmt, 13, SQL_C_SSHORT,&(x->run_status), 0, &temp_int);
						flag=1;
                        retcode = SQLFetch(hstmt);						
                     }//exec
                }//alloc stmt
				x->mathine=mathine;
                SQLFreeStmt(hstmt, SQL_DROP);
			return flag;
};			

/*
	run_flag 的含义：
	   -1---正在运行
		0----停止
		1----开始运行
		2----模拟
		3----在线修改
		4----在线修改确认 后我run_flag==1
		10---Stop any way then VBrun_falg=0
*/
 int WINAPI  write_run(struct run_table *x,int run_flag)
  {
	 
		switch(run_flag)
		{
		case 10:
			can_run_flag=0;
			Pai_Fan_Run_Flag=0;
			Prev_Run_Flag=0;
			break;				
		case 4:			
			can_run_flag=Prev_Run_Flag;		
			Prev_Run_Flag=0;
			break;
		case 0:
			can_run_flag=0;
			Pai_Fan_Run_Flag=0;
			break;
		case 3:					//in Vb is real change by C1.c can_read
			Prev_Run_Flag=can_run_flag;
			return;
	
		}
	 if (can_run_flag!=3)
	 {
			send_run_table.now_che=x->now_che;	/*第几车*/
			send_run_table.all_time=x->all_time;	  /*总工作时间 */
			send_run_table.all_che_time=x->all_che_time;	/*总一车时间 */
			send_run_table.duan_time=x->duan_time; /*段时间 */
	 }
      //  memcpy((void *)&send_run_table,x, sizeof(struct run_table));	 
	else
		{
			send_run_table.now_che=x->now_che;	/*第几车*/
			send_run_table.all_time=x->all_time;	  /*总工作时间 */
			send_run_table.all_che_time=x->all_che_time;	/*总一车时间 */
			send_run_table.duan_time=x->duan_time; /*段时间 */
		}

/*	 else
	 {
		lstrcpy(x->name,send_run_table.name);
		lstrcpy(x->_number,send_run_table._number);
//		x->total_che=send_run_table.total_che;
		send_run_table.now_che=x->now_che;	//第几车
//		x->all_duan=send_run_table.all_duan;// 总段 
//		x->ban=send_run_table.ban;		//班  
		send_run_table.mathine=x->mathine;  //机号 
		send_run_table.all_time=x->all_time;	 //*总工作时间 
		send_run_table.all_che_time=x->all_che_time;	//总一车时间 
		send_run_table.duan_time=x->duan_time; //*段时间 
	 }
*/
    return 1;
}
 /***************
	以下为风送系统
 ***************/
int WINAPI read_Fsml_table(struct Fs_ml_table *x,int *flag,int *number)
{
    char  select_str[400];
	HSTMT hstmt; 
	RETCODE retcode;
	*number=0;
	*flag=Pai_Fan_Run_Flag;


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
            retcode = SQLConnect(ljxt_hdbc, (unsigned char *) "ljxt",SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
			if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
			{
				write_error(ODBC_ERROR,"没设置ljxt ODBC");				
				SQLFreeEnv(ljxt_henv);
				ljxt_henv=NULL;
				ljxt_mdb_open_flag=0;
				return FALSE;
			}			
		}//end SQLALLOCConnect
		else
			return FALSE;	
	}//end SQLALLocEnv
		else
			return FALSE;
};//end if comm_mdb_open_flag
		if(Pai_Fan_Run_Flag!=STOP)
		{
			int i=0;
                retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS||retcode==SQL_SUCCESS_WITH_INFO)
                {
					
					
					lstrcpy(select_str," select * from  Fen_Pei_table where mathine<3 order by sort_idx");

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    while (retcode == SQL_SUCCESS||retcode==SQL_SUCCESS_WITH_INFO)
                    {
                        
						DWORD   temp_int;						
						SQLBindCol(hstmt, 1, SQL_C_SLONG,&(x[i].Fm_ch), 0, &temp_int);
						SQLBindCol(hstmt, 2, SQL_C_SLONG,&(x[i].Fm_jzh), 0, &temp_int);
						SQLBindCol(hstmt, 3, SQL_C_SLONG,&(x[i].Fm_dh), 0, &temp_int);
						SQLBindCol(hstmt, 4, SQL_C_SLONG,&(x[i].Fm_pzh), 0, &temp_int);
						SQLBindCol(hstmt, 5, SQL_C_CHAR,x[i].Fm_clbh, 8, &temp_int);
						SQLBindCol(hstmt, 6, SQL_C_CHAR,x[i].Fm_clname, 18, &temp_int);
						SQLBindCol(hstmt, 7, SQL_C_SSHORT,&(x[i].Fm_sgs), 0, &temp_int);
						SQLBindCol(hstmt, 8, SQL_C_SSHORT,&(x[i].Fm_lsmgs), 0, &temp_int);
						SQLBindCol(hstmt, 9, SQL_C_SSHORT,&(x[i].Fm_bl), 0, &temp_int);						
                        retcode = SQLFetch(hstmt);						
						i++;
                     }//exec end where
                }//alloc stmt
				if(i>0) i--;
				*number=i;
                SQLFreeStmt(hstmt, SQL_DROP);
		}
		else
				*number=0;
			*flag=Pai_Fan_Run_Flag;
			return 1;

}
int WINAPI write_Fsml_table(struct Fs_ml_table *x,int flag,int number)
{
		Pai_Fan_Run_Flag=flag;
		return 1;
}
int WINAPI write_Fs_command(int flag)
{
		Pai_Fan_Run_Flag=flag;
		return 1;
}
int WINAPI read_Fs_command(int *flag)
{
		*flag=Pai_Fan_Run_Flag;
		return 1;
}

int WINAPI read_Fssb_table(int mathine,struct Fs_sb_table x[])
{
    char  select_str[400];
	HSTMT hstmt; 
	RETCODE retcode;
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

                retcode = SQLAllocStmt(ljxt_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					char buffer[20];
					int i=0;
					lstrcpy(select_str," select dou,kind_code,single_weight,work_press,set_low_press, \
set_high_press,congestion_press,max_sendtime,clear_time,stop_use,blxs,temp from  Fen_Send_Table \
where  mathine=");
					_itoa(mathine,buffer, 10 );
					lstrcat(select_str,buffer);
					x[0].Fsb_jzh=0;
                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    while (retcode==SQL_SUCCESS||retcode==SQL_SUCCESS_WITH_INFO)
                    {
                        
						DWORD   temp_int;				
						
						SQLBindCol(hstmt, 1, SQL_C_SLONG,&(x[i].Fsb_dh), 0, &temp_int);
						SQLBindCol(hstmt, 2, SQL_C_SLONG,&(x[i].Fsb_pzh), 18, &temp_int);
						SQLBindCol(hstmt, 3, SQL_C_FLOAT,&(x[i].Fsb_dwight), 0, &temp_int);
						SQLBindCol(hstmt, 4, SQL_C_FLOAT,&(x[i].Fsb_gzyl), 0, &temp_int);
						SQLBindCol(hstmt, 5, SQL_C_FLOAT,&(x[i].Fsb_zysd), 0, &temp_int);
						SQLBindCol(hstmt, 6, SQL_C_FLOAT,&(x[i].Fsb_dysd), 0, &temp_int);
						SQLBindCol(hstmt, 7, SQL_C_FLOAT,&(x[i].Fsb_dsgy), 0, &temp_int);
						SQLBindCol(hstmt, 8, SQL_C_SSHORT,&(x[i].Fsb_mtime), 0, &temp_int);
						SQLBindCol(hstmt, 9, SQL_C_SSHORT,&(x[i].Fsb_qtime), 0, &temp_int);
						SQLBindCol(hstmt, 10, SQL_C_SSHORT,&(x[i].Fsb_ty), 0, &temp_int);
						SQLBindCol(hstmt, 11, SQL_C_FLOAT,&(x[i].Fsb_blxs), 0, &temp_int);
						SQLBindCol(hstmt, 12, SQL_C_SSHORT,&(x[i].Fsb_byl), 0, &temp_int);
                        retcode = SQLFetch(hstmt);						
						if(retcode==SQL_SUCCESS||retcode==SQL_SUCCESS_WITH_INFO)
							x[0].Fsb_jzh++;
						i++;
                     }//exec
                }//alloc stmt
                SQLFreeStmt(hstmt, SQL_DROP);
				return 1;
}
int WINAPI read_fen_ban(struct ban_table *x)
{
    char  select_str[400];
	HSTMT hstmt; 
	RETCODE retcode;
	HENV fen_henv=NULL;
	HDBC fen_hdbc=NULL; 	

	int i=0;
	retcode = SQLAllocEnv(&fen_henv); /* Environment handle */ 
	if (retcode != SQL_SUCCESS) 
	{
				write_error(ODBC_ERROR,"ljxt ODBC Env 失败 ");				
				SQLFreeEnv(fen_henv);
				return FALSE;
	}
        retcode = SQLAllocConnect(fen_henv, &fen_hdbc); /* Connection handle */
        if (retcode != SQL_SUCCESS)
		{
				write_error(ODBC_ERROR,"ljxt ODBC connet 失败");				
				SQLFreeEnv(fen_henv);
				return FALSE;
		}
        SQLSetConnectOption(fen_hdbc, SQL_LOGIN_TIMEOUT, 5);
        retcode = SQLConnect(fen_hdbc, (unsigned char *) "ljxt", SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
		if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
		{
				write_error(ODBC_ERROR,"设置ljxt ODBC 验证失败");				
				SQLFreeEnv(fen_henv);
				return FALSE;
		}			


                retcode = SQLAllocStmt(fen_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					
					lstrcpy(select_str,"select [_number],name,ban,che,mathine  from run_table ");
                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
					x[0].total_che=0;
                    while (retcode==SQL_SUCCESS||retcode==SQL_SUCCESS_WITH_INFO)
                    {
                        
						DWORD   temp_int;										
						SQLBindCol(hstmt, 1, SQL_C_CHAR,x[i]._number, 8, &temp_int);
						SQLBindCol(hstmt, 2, SQL_C_CHAR,x[i].name, 18, &temp_int);
						SQLBindCol(hstmt, 3, SQL_C_SSHORT,&(x[i].ban), 0, &temp_int);
						SQLBindCol(hstmt, 4, SQL_C_SSHORT,&(x[i]._che), 0, &temp_int);
						SQLBindCol(hstmt, 5, SQL_C_SSHORT,&(x[i].mathine), 0, &temp_int);
                        retcode = SQLFetch(hstmt);						
						if(retcode==SQL_SUCCESS||retcode==SQL_SUCCESS_WITH_INFO)
						{
								x[0].total_che+=x[i]._che;
								i++;
						}
						if (i>39) break;						
                     }//exec
					  x[i]._number[0]='\0';
					  x[i].name[0]='\0';
					  x[i].ban=0;	
                }//alloc stmt


        SQLFreeStmt(hstmt, SQL_DROP);

		SQLDisconnect(fen_hdbc);
        SQLFreeConnect(fen_hdbc);        
        SQLFreeEnv(fen_henv);    
		return 1;
}
int WINAPI  read_grap(char *name)
{
	lstrcpy(name,grap_name);
	return 1;
}
int WINAPI  write_grap(BSTR name)
{
	lstrcpy(grap_name,name);
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
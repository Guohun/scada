
#include <windows.h>
#include <stdio.h>
#include <time.h>
  #include "SQL.H"
  #include "SQLext.H"

#include <malloc.h>
#include "dll.h"
#include "temp.h"
#include <memory.h>
#include "api232-c.h"
	#pragma     data_seg("sdata")
      struct power_tempro   work_power={0,-1,598,0,0,0,0};
      struct elect_send_type   elect_input={0,0,{0,0,0},{0,0,0}};
        struct elect_send_type   elect_output={0,0,{0,0,0},{0,0,0}};
       struct  G_error_num G_error={0,0};
       struct dou_type GT_dou={0,0,0,0,0};
       struct dou_type GY_dou={0,0,0,0,0};
       struct dou_type GR_dou={0,0,0,0,0};
       struct ban_type G_ban={0,{0.0f},{0.0f},0,0,0,0};
       CRITICAL_SECTION m_csMyLock={0};
       HANDLE  my_envent=NULL;
       HANDLE  main_thread=NULL;
       struct run_table    send_run_table={0,0,0,0,0,0,0,0,"",""};
       struct Produce_type  send_pei_fan={0,0,0,0,0,0,0};
	   struct make_index  Static_make_index={0,0,0};
       struct  make_table  Static_make_table={0,0,0};

       int  control_can_read=CAN_READ-1;   //初始化
       int  can_run_flag=CAN_RUN-1;
       int  g_mathine=0;
        char  duan_begin_time[20]="";          //段开始时间 由write_power ing定
        char  Grap_File_Name[20]="";
        int   FONT_SIZE=0;
        HFONT  hNewf=NULL;
		int  port_array[8]={0,0,0};
		int Vb_key_code=0;
		HENV elec_henv=NULL;
		HDBC elec_hdbc=NULL; 
		int  elec_mdb_flag=0;
		extern HDC  hdc,memdc;
		
	#pragma data_seg()

HINSTANCE   MyInstance;

extern HBITMAP  hBit;
BOOL  WINAPI  DLLMain(HINSTANCE  hModule, DWORD  fdwReason, LPVOID   lpReserved)
{
	MyInstance =hModule;

	switch(fdwReason)
	{
	case DLL_PROCESS_ATTACH:
	//		MessageBox(0,"asdf","adsfa",MB_OK);        
		 break; 
	case DLL_THREAD_ATTACH:				
			//MessageBox(0,"threadasdf","adsfa",MB_OK);        
			break;
	case  DLL_THREAD_DETACH:
			//	MessageBox(0,"free","free",MB_OK);        
		return 0;
		
	case  DLL_PROCESS_DETACH:
			//	MessageBox(0,"free","free",MB_OK);        
		return 0;
	}
    
return TRUE;
}
int WINAPI  Init_P(void)
{
  InitializeCriticalSection(&m_csMyLock);
	CreateEvent(NULL,FALSE,FALSE,"ABC");
//	main_thread=GetCurrentThread();
	return 1;
}

int WINAPI  Init_V(void)
{
      DeleteCriticalSection(&m_csMyLock);
		CloseHandle(my_envent);
		return 1;
}


int WINAPI  P_DO(void)
{
//			ResetEvent(my_envent);
//        EnterCriticalSection(&m_csMyLock);
//		
//		WaitForSingleObject(my_envent,INFINITE);
}
int WINAPI  V_DO(void)
{
       //LeaveCriticalSection(&m_csMyLock);
	//SetEvent(my_envent);
}

NOMANGLE BSTR  CCONV UpperCaseByRef(BSTR *pBSTROriginal)
{
    BSTR BSTRUpperCase;
    int i;
    int cbOriginalLen;
    BSTR strSrcByRef, strDst;

    #if !defined(_WIN32)
        cbOriginalLen = SysStringLen(*pBSTROriginal);
    #else
        cbOriginalLen = SysStringByteLen(*pBSTROriginal);
    #endif

    BSTRUpperCase = SysAllocStringLen(NULL, cbOriginalLen);

    strSrcByRef = (BSTR)*pBSTROriginal;
    strDst = (BSTR)BSTRUpperCase;

    for(i=0; i<=cbOriginalLen; i++)
        *strDst++ = toupper(*strSrcByRef++);

    SysReAllocString (pBSTROriginal, (BSTR)"Good Bye");
    return BSTRUpperCase;
}

NOMANGLE BSTR CCONV UpperCaseByVal(BSTR BSTROriginal)
{
    BSTR BSTRUpperCase;
    int i;
    int cbOriginalLen;
    BSTR strSrcByVal, strDst;

    #if !defined(_WIN32)
        cbOriginalLen = SysStringLen(BSTROriginal);
    #else
        cbOriginalLen = SysStringByteLen(BSTROriginal);
    #endif

    BSTRUpperCase = SysAllocStringLen(NULL, cbOriginalLen);

    strSrcByVal = (BSTR)BSTROriginal;
    strDst = (BSTR)BSTRUpperCase;

    for(i=0; i<=cbOriginalLen; i++)
        *strDst++ = toupper(*strSrcByVal++);

    SysReAllocString (&BSTROriginal, (BSTR)"Good Bye");

    return BSTRUpperCase;
}
// control->dll
 int WINAPI  write_power(int duan_hao,float power,int tempro,int flag)
{
	if (can_run_flag==STOP_DOING)
	{	
		work_power.flag=-1;
         work_power.ptr=0;
		 duan_hao=0;
		return  STOP_DOING;
	}

     if (flag==-1)//||(duan_hao!=work_power.duan_hao))
     {
		 work_power.tempro[work_power.ptr]=tempro;
	     work_power.power[work_power.ptr]=power;	 
		 work_power.ptr=0;
		 duan_hao=0;
		 return 0;
	 }
/*         work_power.ptr=0;
		 work_power.flag=flag;
	     work_power.tempro[work_power.ptr]=tempro;
		 work_power.power[work_power.ptr]=power;
		 if(flag==-1)  return STOP_DOING;
		 work_power.flag=flag;
		 work_power.duan_hao=duan_hao;
		 return 1;
     }
     else
		 work_power.ptr++;
  */   
     work_power.tempro[work_power.ptr]=tempro;
     work_power.power[work_power.ptr]=power;	 
     work_power.flag=flag;
     work_power.duan_hao=duan_hao;
	 work_power.ptr++;
	 if (work_power.ptr>work_power.total)
			work_power.ptr=0;

    return 1;
}
//dll->vb
NOMANGLE int CCONV read_power(int *duan_hao,float *power,int *tempro,int *flag)
{
	if (can_run_flag==STOP_DOING)
	{		
		*power=0;
		*tempro=0;
		return  STOP_DOING;
	}
	if(work_power.ptr==0)
	{
	
//		*power=0;
//		*tempro=0;
		return 1;
	}
     *flag=work_power.flag;
     *duan_hao=work_power.duan_hao;
     *tempro=work_power.tempro[work_power.ptr-1];
     *power=work_power.power[work_power.ptr-1];
    return 1;
}
//dll->vb
/*
  flag  =0  not work	//dou
		=1   work
		=-1   error
  retrun  0  not work  // mathine
		  1   work
*/


int WINAPI  read_select_ban(int mathine,char *temp)
{
	return 1;
}

int WINAPI  write_select_ban(int mathine,char *temp)
{
	return 1;
}


int WINAPI   write_ban(int mathine,int name,float far * max,float far * weight, int flag)
{
	int i;

	for(i=0;i<G_ban.length;i++)	
		if (G_ban.mathine[i]==mathine&&G_ban.name[i]==name)
			{
				G_ban.weight[i]=*weight;
				G_ban.max[i]=*max;				
				G_ban.flag[i]=flag;				
				return 1;
			}
	if (i>19)
		return -1;
	G_ban.name[i]=name;
	G_ban.mathine[i]=mathine;
	G_ban.max[i]=*max;
	G_ban.weight[i]=*weight;
	G_ban.length=i+1;	
	G_ban.flag[i]=flag;				
	return 1;	
	
}
NOMANGLE int CCONV  read_ban(short mathine,short name,float far *max,float far *weight,int *flag)
{

	int i;

	for(i=0;i<G_ban.length;i++)	
		if (G_ban.mathine[i]==mathine&&G_ban.name[i]==name)
			{	
				*weight=G_ban.weight[i];
				*max=G_ban.max[i];
				*flag=G_ban.flag[i];
				return 1;
			}
	*max=0.0f;
	*weight=0.0f;
	*flag=1;
	return 0;	

}

NOMANGLE int  CCONV   read_Tan_dou(short mathine,short dou_hao,short far *flag)
{
	int i;

//	if (can_run_flag==STOP_DOING)
//		return  STOP_DOING;
    for(i=0;i<GT_dou.length;i++)
        if (GT_dou.mathine[i]==mathine&&GT_dou.name[i]==dou_hao)
			{		
                *flag=GT_dou.flag[i];
				return 1;
			}
	*flag=0;

	return 1;	

}


int WINAPI  write_ri_chu_dou(int mathine,int  dou_hao,int  flag)
{
	int i;

    for(i=0;i<GR_dou.length;i++)
                GR_dou.flag[i]=0;
    for(i=0;i<GR_dou.length;i++)
        if (GR_dou.mathine[i]==mathine&&GR_dou.name[i]==dou_hao)
			{		
                GR_dou.flag[i]=flag;
				return 1;
			}


  if(i>39)
	  return -1;
        GR_dou.name[i]=dou_hao;
        GR_dou.mathine[i]=mathine;
        GR_dou.flag[i]=flag;
        GR_dou.length=i+1;
	return 1;		


}
NOMANGLE int  CCONV read_ri_chu_dou(short mathine,short  dou_hao,short far  *flag)
{
		int i;

//	if (can_run_flag==STOP_DOING)
//		return  STOP_DOING;
    for(i=0;i<GR_dou.length;i++)
        if (GR_dou.mathine[i]==mathine&&GR_dou.name[i]==dou_hao)
			{		
                *flag=GR_dou.flag[i];
				return 1;
			}
	*flag=0;

	return 1;
}
int WINAPI   write_Tan_dou(int mathine,int dou_hao,int flag)
{
	int i;

//	if (can_run_flag==STOP_DOING)
//		return  STOP_DOING;
    for(i=0;i<GT_dou.length;i++)
                GT_dou.flag[i]=0;
    for(i=0;i<GT_dou.length;i++)
        if (GT_dou.mathine[i]==mathine&&GT_dou.name[i]==dou_hao)
			{		
                GT_dou.flag[i]=flag;
				return 1;
			}


  if(i>39)
	  return -1;
        GT_dou.name[i]=dou_hao;
        GT_dou.mathine[i]=mathine;
        GT_dou.flag[i]=flag;
        GT_dou.length=i+1;
	return 1;		
}


NOMANGLE int  CCONV  read_You_dou(short mathine,short dou_hao,int *flag)
{
	int i;

//	if (can_run_flag==STOP_DOING)
//		return  STOP_DOING;
    for(i=0;i<GY_dou.length;i++)
        if (GY_dou.mathine[i]==mathine&&GY_dou.name[i]==dou_hao)
			{		
                *flag=GY_dou.flag[i];
				return 1;
			}
	*flag=0;

	return 1;	

}



int WINAPI   write_You_dou(int mathine,int dou_hao,int flag)
{
	int i;

//	if (can_run_flag==STOP_DOING)
//		return  STOP_DOING;
    for(i=0;i<GY_dou.length;i++)
                GY_dou.flag[i]=0;
    for(i=0;i<GY_dou.length;i++)
        if (GY_dou.mathine[i]==mathine&&GY_dou.name[i]==dou_hao)
			{		
                GY_dou.flag[i]=flag;
				return 1;
			}


  if(i>39)
	  return -1;
        GY_dou.name[i]=dou_hao;
        GY_dou.mathine[i]=mathine;
        GY_dou.flag[i]=flag;
        GY_dou.length=i+1;
	return 1;		
}
int WINAPI   init_elec_input(struct elect_send_type *x)
{
//	
	int i;
//	if (can_run_flag==STOP_DOING)
//	{		
//		return  STOP_DOING;
//	}     
	x->flag=elect_input.flag;
//	if (elect_input.flag==0)  return 1;
	
//	memcpy((void *)x,(void *)&elect_input,sizeof(struct elect_send_type) );			
//	return 1;

	x->length=elect_input.length;
	for(i=0;i<=x->length;i++)
	{
//		for(j=0;j<7;j++)
    	x->note_name[i]=elect_input.note_name[i];
		x->data[i]=elect_input.data[i];
	}
	elect_input.flag=0;
	return 1;
}


 int WINAPI   init_elec_output(struct elect_send_type *x)
{

	int i;
//	if (can_run_flag==STOP_DOING)			
//		return  STOP_DOING;
	x->flag=elect_output.flag;
//	if (elect_output.flag==0)  return 1;
	
	x->length=elect_output.length;
	for(i=0;i<=x->length;i++)
	{

		x->note_name[i]=elect_output.note_name[i];
		x->data[i]=elect_output.data[i];
	}
//	elect_output.flag=0;
	return 1;
}


//dll--->control  or  dll-->vb
NOMANGLE  int  CCONV read_elec_input(struct elect_send_type *x)
{
	int i;
/*	if (can_run_flag==STOP_DOING)
	{		
		return  STOP_DOING;
	}     
*/
	x->flag=elect_input.flag;
//    if (elect_input.flag==0)  return 1;	
	x->length=elect_input.length;
	for(i=0;i<=elect_input.length;i++)
	{
		x->note_name[i]=elect_input.note_name[i];
		x->data[i]=elect_input.data[i];
//		MessageBox(0,"read_input","",0);
	}
//	elect_input.flag=0;
	return 1;
}
//dll-->control  or dll--->vb
NOMANGLE int   CCONV read_elec_output(struct elect_send_type *x)
{

	int i;
	x->flag=elect_output.flag;
//	if     (elect_output.flag==0) return  1;
	
/*	elect_output.data[0]=1;
	elect_output.data[1]=2;
	elect_output.data[2]=3;
	elect_output.data[4]=4;
*/
	
	x->length=elect_output.length;
	for(i=0;i<=x->length;i++)
	{
		x->data[i]=elect_output.data[i];
		x->note_name[i]=elect_output.note_name[i];
		
	}
//	elect_output.flag=0;
	return 1;
}

//control--->dll
 int WINAPI   write_elec_input(struct elect_send_type *x)
{
	int i;
	switch(x->flag)
	{
	case 1:						//first_time doing
			elect_input.flag=1;
	case 2:
			elect_input.flag=2;	//vb-->dll
	case 3:						//vb control-->dll
			elect_input.flag=0;
	default:
			elect_input.flag=0;
	};
	elect_input.length=x->length;
	for(i=0;i<=x->length;i++)
	{
		elect_input.note_name[i]=x->note_name[i];
		elect_input.data[i]=x->data[i];
	}
//	memcpy((void *)&elect_input,(void *)x,sizeof(struct elect_send_type) );	
	return 1;
}
//control-->dll
int WINAPI write_elec_output(struct elect_send_type *x)
{
	int i;
	switch(x->flag)
	{
	case 1:						//first_time doing
			elect_output.flag=1;
	case 2:
			elect_output.flag=2;	//vb-->dll
			return ;
	case 3:						//vb control-->dll
			elect_output.flag=0;
			return ;
	default:
			elect_output.flag=0;
	};

	elect_output.length=x->length;
	for(i=0;i<=elect_output.length;i++)
	{
		elect_output.note_name[i]=x->note_name[i];
		elect_output.data[i]=x->data[i];
	}
	return 1;
}
//control-->dll
/*
    if error_num=-1  then
    add  end
    return  error_num
*/
int WINAPI  write_error(int  error_num,char *prompt)
{
	int i;
	for(i=0;i<G_error.length;i++)
	{
		if (G_error.data[i]==error_num)		
				return WRITE_DOING;		
	}
    lstrcpy(G_error.prompt[i],prompt);
	G_error.data[i]=error_num;
	G_error.length++;
	return  1;
}
//dll->control
int WINAPI read_error(short  error_num)
{
	int i;
	for(i=0;i<G_error.length;i++)
	{
		if (G_error.data[i]==error_num)		
				return WRITE_DOING;		
	}

	return 0;
}
//control-->control
int WINAPI kill_error(int error_num)
{
	int i;
	for(i=0;i<G_error.length;i++)
	{
		if (G_error.data[i]==error_num)		
				break;
	}
    if (i==G_error.length)
			return  WRITE_DOING;
	for(;i<G_error.length;i++)
			G_error.data[i]=G_error.data[i+1];
	G_error.length--;
	return 0;
}

int WINAPI get_error(struct G_error_num * error_num)
{
    int i;
    error_num->length=G_error.length;
     for(i=0;i<G_error.length;i++)
     {
       error_num->data[i]=G_error.data[i];
       lstrcpy(error_num->prompt[i],G_error.prompt[i]);
    }
	 if(G_error.length>0)
		return 1;
	 else
		return 0;
}





int WINAPI  Z_DLL_Close_All()
{
	int i=0;
	for(i=0;i<8;i++)
		if(port_array[i]!=0)
	{
			sio_close(port_array[i]);
			port_array[i]=0;
	}
	else
		break;
	if(comm_mdb_open_flag)
	{
		comm_mdb_open_flag=0;
		SQLDisconnect(comm_hdbc);
        SQLFreeConnect(comm_hdbc);        
        SQLFreeEnv(comm_henv);    
	}
	if(make_mdb_open_flag)
	{
		make_mdb_open_flag=0;
		SQLDisconnect(make_hdbc);
        SQLFreeConnect(make_hdbc);        
        SQLFreeEnv(make_henv);    
	}
	if(ljxt_mdb_open_flag)
	{
		ljxt_mdb_open_flag=0;
//		SQLDisconnect(ljxt_hdbc);
//        SQLFreeConnect(ljxt_hdbc);        
//        SQLFreeEnv(ljxt_henv);    
	}
	if(elec_mdb_flag)
	{
		elec_mdb_flag=0;
		SQLDisconnect(elec_hdbc);
        SQLFreeConnect(elec_hdbc);        
        SQLFreeEnv(elec_henv);    

	}
	DeleteDC(memdc);
	ReleaseDC(0,hdc);
	return TRUE;
}
int   port,bound,data,stop,parity;

int WINAPI  Sio_Open_All(int mathine)
{
    char  select_str[400];

	int  i=0;
	
	HSTMT hstmt; 
	RETCODE retcode;
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
				i=0;
                retcode = SQLAllocStmt(comm_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {
					char buffer[20];
					lstrcpy(select_str," select port,bound,data,stop,parity from  comm_table \
where  mathine=");
					_itoa(mathine,buffer, 10 );
					lstrcat(select_str,buffer);

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    while(retcode == SQL_SUCCESS)
                    {                        
						DWORD   temp_int;						
						SQLBindCol(hstmt, 1, SQL_C_STINYINT,&port, 0, &temp_int);
						SQLBindCol(hstmt, 2, SQL_C_STINYINT,&bound, 0, &temp_int);
						SQLBindCol(hstmt, 3, SQL_C_STINYINT,&data, 0, &temp_int);
						SQLBindCol(hstmt, 4, SQL_C_STINYINT,&stop, 0, &temp_int);
						SQLBindCol(hstmt, 5, SQL_C_STINYINT,&parity, 0, &temp_int);

                        retcode = SQLFetch(hstmt);						
						if(retcode!= SQL_SUCCESS) goto error_do;
						if (sio_open(port)<0)	goto   error_do;
						if(sio_ioctl(port,bound, data| stop| parity)<0) goto error_do;          
						port_array[i++]=port;
						if(i>8) break;
                     }//exec
                }//alloc stmt
                SQLFreeStmt(hstmt, SQL_DROP);

	return TRUE;
error_do:
            SQLFreeStmt(hstmt, SQL_DROP);
			return FALSE;
}

int WINAPI Get_Elec_Check_Table(struct	elec_in_table x[],int *n)
{
	int i=0;
    char  select_str[400];

	HSTMT hstmt; 
	RETCODE retcode;

if (elec_mdb_flag==0)
{
	elec_mdb_flag=1;
	retcode = SQLAllocEnv(&elec_henv); /* Environment handle */ 
	if (retcode == SQL_SUCCESS) 
	{ 
        retcode = SQLAllocConnect(elec_henv, &elec_hdbc); /* Connection handle */
        if (retcode == SQL_SUCCESS)
        {
            SQLSetConnectOption(elec_hdbc, SQL_LOGIN_TIMEOUT, 5);
            retcode = SQLConnect(elec_hdbc, (unsigned char *) "ljxt", SQL_NTS, (unsigned char *)"", SQL_NTS, (unsigned char *)"", SQL_NTS);
			if (retcode != SQL_SUCCESS && retcode != SQL_SUCCESS_WITH_INFO)
			{
				write_error(ODBC_ERROR,"没设置ljxt ODBC");				
				SQLFreeEnv(elec_henv);
				elec_henv=NULL;
				elec_mdb_flag=0;
				return FALSE;
			}			
		}//end SQLALLOCConnect
	}//end SQLALLocEnv
};//end if comm_mdb_open_flag
				i=0;
                retcode = SQLAllocStmt(elec_hdbc, &hstmt); /* Statement handle */
                if (retcode == SQL_SUCCESS)
                {				
lstrcpy(select_str," select name,str,run,run_set,check,check_set, run_value,check_value,check_time,use_flag from  elec_in_table ");

                    retcode=SQLExecDirect(hstmt,select_str,SQL_NTS);
                    while(retcode == SQL_SUCCESS)
                    {                        
						DWORD   temp_int;						
						if(i>=*n) break;
						SQLBindCol(hstmt, 1, SQL_C_SSHORT,&(x[i].name), 0, &temp_int);
						SQLBindCol(hstmt, 2, SQL_C_CHAR,x[i].str, 4, &temp_int);
						SQLBindCol(hstmt, 3, SQL_C_SSHORT,&(x[i].run), 0, &temp_int);
						SQLBindCol(hstmt, 4, SQL_C_SSHORT,&(x[i].run_set), 0, &temp_int);
						SQLBindCol(hstmt, 5, SQL_C_SSHORT,&(x[i].check), 0, &temp_int);
						SQLBindCol(hstmt, 6, SQL_C_SSHORT,&(x[i].check_set), 0, &temp_int);
						SQLBindCol(hstmt, 7, SQL_C_SSHORT,&(x[i].run_value), 0, &temp_int);
						SQLBindCol(hstmt, 8, SQL_C_SSHORT,&(x[i].check_value), 0, &temp_int);
						SQLBindCol(hstmt, 9, SQL_C_SSHORT,&(x[i].check_time), 0, &temp_int);
						SQLBindCol(hstmt, 10, SQL_C_SSHORT,&(x[i].use_flag), 0, &temp_int);
                        retcode = SQLFetch(hstmt);						
						if(retcode!= SQL_SUCCESS) break;
						i++;						
                     }//exec
//					SQLTransact(ljxt_henv, ljxt_hdbc, SQL_COMMIT);
                }//alloc stmt
                SQLFreeStmt(hstmt, SQL_DROP);
				*n=i;	
		return TRUE;
}

int  Get_Date(char  *buffer )
{
   char dbuffer [13];   
   _strdate( dbuffer );
   lstrcpy( buffer,dbuffer);

	return 1;
}
int  Get_Time(char  *buffer)
{
   char tbuffer [9];
   _strtime( tbuffer );
   lstrcpy( buffer, tbuffer );
   return 1;
}

//Output

int WINAPI Set_Light(int elec_name,int data)
{
	int k;
           for(k=0;k<=elect_input.length;k++)
                if (elec_name==elect_input.note_name[k])     //close upper
				  {
                            elect_input.data[k]=data;
//							elect_input.flag=2;
                                break;
				  }
			
			return 1;
}
int WINAPI Set_Turn(int elec_name,int data)
{
	int k;
           for(k=0;k<=elect_output.length;k++)
                if (elec_name==elect_output.note_name[k])     //close upper
				  {
                            elect_output.data[k]=data;
//							elect_output.flag=2;	
                            break;
				  }               
		    return 1;
}

int WINAPI Read_VbKey()
{		
			int temp;
			temp=Vb_key_code;
			Vb_key_code=0;
		    return temp;
}
int WINAPI Write_VbKey(int key_code)
{

		     Vb_key_code=key_code;
			 return 0;
}

float Round(float x,int number)
{
	int i;
	long  temp;
	for(i=0;i<number;i++)
	 x=x*10;
	 x=x*10;

	temp=(long)x;
	if( temp%10>5)
		temp+=10;
	temp=temp/10;
	x=0;
	for(i=0;i<number;i++)
		{
			x=(x+((float)(temp%10)))/10;
			temp=temp/10;
		}
	return (float)(temp+x);
}

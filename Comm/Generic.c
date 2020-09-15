/*****************************************************
	��������ϵͳ:�������Գ���
	�����      :������,�����
	����	     :1998.1.12
	������Ҫ��  :Vitual C++ ��Bland C++
	����Ҫ��    :��������Ѱ�װ
	��������    :���û���ͨѶ ,������,���߳�,�䷽���� ,����ͳ���ӳ���
		   	���Զ��豸����.������Ĳ���������ͨ��.
*****************************************************/
#include <windows.h>
#include <stdio.h>
#include <malloc.h>
#include "Resrc1.h"
#include "genwin.h"
#include "dll.h"
#include "api232-c.h"

#define TEST_LEN    (1024 * 10)		//����Buffer ����
#define TEST_PORT   2			//���� �� 2
int  WINAPI can_read(BSTR * pei_number,BSTR *name,short ban,char total_che,short flag);
HANDLE   hThread;				//���߳�   
HANDLE    chdthd[10];			//�������߳�   ���10�� 
HWND      hwnd;				//���߳� hwnd
int     total_thread=0;			//��ǰ�������̸߳���
int     run_flag=STOP_DOING;      		//���б�־

/*****************************************************
	Windows ��������  �ɲο�һ��Windows������
*****************************************************/
HANDLE          MyInstance;
static char     GenericClass[32]=" ���Գ���";
static BOOL FirstInstance( HANDLE );
static BOOL AnyInstance( HANDLE, int, LPSTR );
long  FAR PASCAL WindowProc( HWND, unsigned, UINT, LONG );
/*****************************************************
	Windows ������  �ɲο�һ��Windows������
*****************************************************/
int PASCAL WinMain( HANDLE this_inst, HANDLE prev_inst, LPSTR cmdline,
                 int cmdshow )
{
    MSG         msg;
    MyInstance = this_inst;
    if( !prev_inst ) {
        if( !FirstInstance( this_inst ) ) return( FALSE );
    }
    if( !AnyInstance( this_inst, cmdshow, cmdline ) ) return( FALSE );

    while( GetMessage( &msg, NULL, NULL, NULL ) ) {

        TranslateMessage( &msg );
        DispatchMessage( &msg );

    }
     return FALSE;
    return( msg.wParam );

} /* WinMain */

/*****************************************************
	дͨѶ�ڲ����߳� 
*****************************************************/
LRESULT WINAPI write_Thread (int x)
{
	     char  temp[TEST_LEN+1];
	     int i=0;
	     HDC   hdc;
	       if (sio_open(TEST_PORT)<0)	//���Ժ��� sio_open		
		  {
			   MessageBox(hwnd,"open  error","open",0);
			   return  0;
		  }
	   sio_ioctl(TEST_PORT,B9600, BIT_8 | STOP_1 | P_NONE);          
	   while (TRUE)
	    {
              wsprintf(temp,"%10d",++i);
              if(sio_write(TEST_PORT, temp, TEST_LEN)<0) 
		{
			MessageBox(hwnd,"write  error","open",0);
			return 0;
		}						
		hdc=GetDC(hwnd);
              wsprintf(temp,"write %10d ",i);
              TextOut(hdc,10,10,temp,40);
		ReleaseDC(hwnd,hdc);
		if(i>4000)
		{											
			hdc=GetDC(hwnd);
			sprintf(temp,"%8s","˯�� 1s");
			TextOut(hdc,10,10,temp,8);
			ReleaseDC(hwnd,hdc);
			Sleep(1000);		//˯��1000ms  ��1S
			i=1;
		}
	}
 return FALSE;
}
/*****************************************************
��ͨѶ�ڲ����߳�   ��PLC����  ��������,��������
*****************************************************/
LRESULT WINAPI read_Thread (int x)
{
    char  temp[40],temp1[40],s1[40],s2[40];
    int i=0,j,k;
    unsigned c1;
    float temp_float,f1;
    static int static_d=0;	
    HDC  hdc;

      sio_close(2);	//Ϊ��ȫ �ȹر� 2��
      Sleep(200);
      if(sio_open(2)<0){
		   MessageBox(hwnd,"open  error","open 2",0);
       	   return  0;
	 }
       sio_ioctl(2,B2400, BIT_7 | STOP_2 | P_EVEN);
       while (TRUE){
                hdc=GetDC(hwnd);
  		  sio_flush(2,0);	
                Sleep(100);
		do{
                   i=sio_getch(2);
    		     Sleep(10);
                }while(i!=2);
                sio_read(2,s1,9);
		  c1=s1[0]&0x007;
		 for(i=0;i<7;i++)  s2[i]=s1[i+3];
 		  k=atoi(s2);
		  switch(c1){
			case 0x01: f1=k*10.0;break;
			case 0x03: f1=k/10.0;break;
			case 0x04: f1=k/100.0;break;
			case 0x05: f1=k/1000.0;break;
			case 0x06: f1=k/10000.0;break;
			};
              if((s1[1]&0x02)!=0) f1=-f1;
              sprintf(temp1,"%8.4f",f1);
              TextOut(hdc,10,200,temp1,8);
 	       temp_float=100;
		//���� д�Ӻ���
//		write_ban(1,1,&temp_float,&f1,2);  		
		if (static_d==0) {
//			write_error(1,"abc");
			static_d=1;
		   }
   		   sio_flush(2,0);
                Sleep(200);
              };
             ReleaseDC(hwnd,hdc);
             if(sio_close(2)<0) {
                       MessageBox(hwnd,"close  error","open",0);
                        return  0;
                 }
}

/*****************************************************
���Զ����ź��߳�   �ӿ��Ƴ����  
*****************************************************/
struct  elect_send_type     elec_input;  //  ���Ե��ź�  �����ź�����
struct  elect_send_type     elec_output;  // ���Ե��ź�  �����ź�����
LRESULT WINAPI Read_Elec_Thread (int x)
{
	char temp[280];
		HDC  hdc;
	 do{
         if (run_flag!=STOP_DOING)
         {			   		   
		int x;
		//���Ƴ����ʼ�� ���ź�
//  	       init_elec_input(&elec_input);
//	       init_elec_output(&elec_output);
		
		   hdc=GetDC(hwnd);
		   sprintf(temp,"�����ź��ܳ���=%d ����ź��ܳ���=%d",
			elec_input.length,elec_output.length);
		   x=strlen(temp);
	          TextOut(hdc,60,120,temp,x);					  
		   ReleaseDC(hwnd,hdc);
		   write_multi_screen(120,60,temp);	//д������
        }
   	Sleep(1000);
 }while(1);
return FALSE;
}

/*****************************************************
���Է���ϵͳ�߳�   �ӿ��Ƴ����  
*****************************************************/
LRESULT WINAPI Write_Elec_Thread (int y)
{

static    float data=0.5;
static	struct Fs_ml_table  fs[80]={0,0,0};
	int flag,number,k,i,x;
static	char  grap_name[40];
static	struct ban_table ban[40];
static	struct Fs_sb_table  fx[50]={0,0,0,0};
    HDC  hdc;
static	char temp[200];
	do{
		Set_Light(1227,1);
		Set_Turn(1227,1);
		for(i=0;i<8;i++)
		{
//		write_Tan_dou(1,i,1);
		Sleep(1000);
		}
		hdc=GetDC(hwnd);
//	    read_grap(grap_name);
	    sprintf(temp,"ͼ�����Ϊ:%s",grap_name);
		x=strlen(temp);
		TextOut(hdc,20,10,temp,x);					  		
		if (strcmp(grap_name,"main_tk")!=0)
		{

			if (run_flag==STOP_DOING)
			{			   		   
		//		read_fen_ban(ban);
		/*		for(i=0;i<40&ban[i]._number[0]!='\0';i++)
				{
				sprintf(temp,"��������Ϊ: ����%s,����%s,��%d,��%d,��%d,�ܳ�%d",ban[i]._number,
					ban[i].name,ban[i].ban,ban[i]._che,ban[i].mathine,ban[0].total_che);
				TextOut(hdc,20,30+i*20,temp,strlen(temp));					  
				}
			*/
			}
			Sleep(6000);
			 ReleaseDC(hwnd,hdc);  		    

			continue;
		}
		Sleep(10);
		hdc=GetDC(hwnd);
//		read_Fssb_table(2,fx);
		   sprintf(temp,"��2 ���� �����豸"
			);
			x=strlen(temp);
	        TextOut(hdc,60,120,temp,x);					  

			for(i=0;i<fx[0].Fsb_jzh;i++)
			{
			    sprintf(temp,"����=%d Ʒ�ֺ�=%d,����=%f,����ѹ��=%f,��ѹ=%f,��ѹ=%f,��ѹ=%f,����=%d,��ɨ=%d,ͣ��=%d,����=%d,����=%d",
				fx[i].Fsb_dh,fx[i].Fsb_pzh,fx[i].Fsb_dwight,fx[i].Fsb_gzyl,fx[i].Fsb_zysd,fx[i].Fsb_dysd,fx[i].Fsb_dsgy,fx[i].Fsb_mtime,
				fx[i].Fsb_qtime,fx[i].Fsb_ty,fx[i].Fsb_blxs,fx[i].Fsb_byl);
			    x=strlen(temp);
				TextOut(hdc,20,140+i*20,temp,x);					  					
			}


//		read_Fsml_table(fs,&flag,&number);
		hdc=GetDC(hwnd);
		   sprintf(temp,"����ϵͳ����=%d ��������=%d",
			number,flag);
	    TextOut(hdc,20,220,temp,strlen(temp));					  

       if(flag!=0)
	   {
		   int send1,time1;
//		   write_Fsml_table(fs,2,number);
//		   write_chans(1,2,3,4);
//		   write_runfch(data+2,data,2*data,data*4,data+3,"����");
		   send1=data+2;
			time1=data+3;
		   sprintf(temp,"����=%4d,1#=%4.6f 2#=%4.6f , ѹ��=%4.6f,ʱ��=%4d",
			send1,data,2*data,data*4,time1);
		    TextOut(hdc,20,240,temp,strlen(temp));					  
		   data+=2.4;
	   }
	   if (data>100)
		   data=0.5;
	   

		read_fen_ban(ban);
		for(i=0;i<40&ban[i]._number[0]!='\0';i++)
		{
		    sprintf(temp,"��������Ϊ: ����%s,����%s,��%d,��%d,��%d,�ܳ�%d",ban[i]._number,
				ban[i].name,ban[i].ban,ban[i]._che,ban[i].mathine,ban[0].total_che);
			TextOut(hdc,20,30+i*20,temp,strlen(temp));					  

		}

	   for(i=0;i<number;i++)
	   {
		   sprintf(temp,"���=%d ����=%d,����=%d,Ʒ�ֺ�=%d,����=%s,����=%s,����=%d,����=%d,����=%d",
			fs[i].Fm_ch,fs[i].Fm_jzh,fs[i].Fm_dh,fs[i].Fm_pzh,fs[i].Fm_clbh,
			fs[i].Fm_clname,fs[i].Fm_sgs,fs[i].Fm_lsmgs,fs[i].Fm_bl);			      
			TextOut(hdc,20,260+20*i,temp,strlen(temp));					  

	   }
		   ReleaseDC(hwnd,hdc);
		Sleep(1000);
		   
    }while(1);

    return FALSE;
}


/**********************************************************
��������ͳ�������߳�   ���������������� ,���򲻻������� ��ʾ�Ͳ���
**********************************************************/
struct make_base_type Make_Temp;	//  ����ÿ�� ��������
struct make_lian_liao Duan_Data;	//  ����ÿ�� ��������
LRESULT WINAPI Read_Run_Thread (int x)
{
struct  run_table run_table_do;
struct  Produce_type p_y;
int i; char temp[280];
    do{
       static int doing_che=0;       
       char   *statu[6]={"ֹͣ", "����","ģ��"};
  //     read_run(&run_table_do, &run_flag);
       if (run_flag!=STOP_DOING)
       {			   	
	   	HDC  hdc;
	       hdc=GetDC(hwnd);
	       sprintf(temp,"�ܳ���=%d,����=%d,�䷽���=%s,�䷽����=%s",run_table_do.total_che,
		run_table_do.mathine,run_table_do._number,run_table_do.name);
	       run_table_do.now_che++;//��ǰ����++				
  	       if(run_table_do.now_che>run_table_do.total_che)//�жϽ������䷽��
              {
                  run_table_do.now_che=0;			//   
    //              write_run(&run_table_do,0);		//����=0  ��ʾ����	
	  //	    write_power(4,50,60,-1);  //write �¶�,����
		    sprintf(temp,"���н���");					
		    x=strlen(temp);  TextOut(hdc,60,40,temp,x);				
		    return FALSE;
                }else
                    write_run(&run_table_do,-1);	//����=-1  ��ʾ����	
	        x=strlen(temp);
               TextOut(hdc,60,40,temp,x);				
	        write_multi_screen(160,60,temp);
               read_produce(&p_y );
		 sprintf(temp,"̼��Ͷ1��=%d,����1Ͷ1=%d,����2Ͷ1=%d ������=%d �ܶ���=%d",
		           p_y.th1_sum,p_y.yz1_sum,p_y.yt1_sum,p_y.sj1_sum,p_y.pei_max);
		 x=strlen(temp); TextOut(hdc,60,260,temp,x);				
	        sprintf(temp,"name=%s,��1����=%f,������=%f,fast_do=%f,drop_do=%f",p_y.yz1[0].name,p_y.yz1[0].weight,
			p_y.yz1[0].gon_cha,p_y.yz1[0].fast_do,p_y.yz1[0].drop_do);
		 x=strlen(temp);   TextOut(hdc,60,180,temp,x);				
	        sprintf(temp,"name=%s,��2����=%f,������=%f,fast_do=%f,drop_do=%f",p_y.yz2[0].name,p_y.yz2[0].weight,
			p_y.yz2[0].gon_cha,p_y.yz2[0].fast_do,p_y.yz2[0].drop_do);
		  x=strlen(temp);TextOut(hdc,60,200,temp,x);				
	         sprintf(temp,"����1=%2d,����2=%2d,̼��1=%2d,̼��2=%2d,��1=%2d,��2=%2d",
			p_y.sj1_sum,p_y.sj2_sum,p_y.th1_sum,p_y.th2_sum,p_y.yz1_sum,p_y.yz2_sum);
	         x=strlen(temp); TextOut(hdc,60,240,temp,x);				
			 if (p_y.yt1_sum>0)
			 {
				sprintf(temp,"name=%s,��21����=%f,������=%f,fast_do=%f,drop_do=%f",p_y.yt1[0].name,p_y.yt1[0].weight,
				p_y.yt1[0].gon_cha,p_y.yt1[0].fast_do,p_y.yt1[0].drop_do);
				x=strlen(temp);   TextOut(hdc,60,260,temp,x);				

			 }
			 if (p_y.yt2_sum>0)
			 {
		        sprintf(temp,"name=%s,��22����=%f,������=%f,fast_do=%f,drop_do=%f",p_y.yt2[0].name,p_y.yt2[0].weight,
				p_y.yt2[0].gon_cha,p_y.yt2[0].fast_do,p_y.yt2[0].drop_do);
				x=strlen(temp);TextOut(hdc,60,280,temp,x);				
			 }

		   ReleaseDC(hwnd,hdc);
		//   write_multi_screen(200,60,temp);
	
		   //����һ����ʼ
	/*�Ȳ���һ�����õ����ж�����*/
	for(i=0;i<p_y.pei_max;i++)
	{	
				int j;
		//		for(j=0;j<30;j++)
		///		write_power(i+1,j+30,(j+30)%40,1);  //write �¶�,����

				Duan_Data.duan_hao=i+1;
				Duan_Data.get_hun_time=p_y.huan[i].lian_time;
				Duan_Data.get_you_time=p_y.huan[i].dou_you_time;
				Duan_Data.get_you2_time=p_y.huan[i].dou_you2_time;
				Duan_Data.get_jiao_time=p_y.huan[i].dou_jiao_time;
				Duan_Data.get_tan_time=p_y.huan[i].dou_tan_time;
				Duan_Data.get_xiao_time=p_y.huan[i].dou_xiao_time;
						
				Duan_Data.next_tempro=p_y.huan[i].tem+60-i;
				Duan_Data.next_power=p_y.huan[i].neg+i*10;
							
				lstrcpy(Duan_Data.duan_begin_time,"11:12:12");
				lstrcpy(Duan_Data.duan_end_time,"11:22:12");
		//		write_make(Duan_Data);
			}
			Make_Temp.pei_max=i;
			MessageBox(hdc,"д������","Test ",MB_OK);
		    memcpy(Make_Temp.th1,p_y.th1,sizeof(p_y.th1));			   
			memcpy(Make_Temp.th2,p_y.th2,sizeof(p_y.th2));			   
		    memcpy(Make_Temp.sj1,p_y.sj1,sizeof(p_y.sj1));			   
			memcpy(Make_Temp.sj2,p_y.sj2,sizeof(p_y.sj2));			   
			memcpy(Make_Temp.yz1,p_y.yz1,sizeof(p_y.yz1));			   
			memcpy(Make_Temp.yz2,p_y.yz2,sizeof(p_y.yz2));			   
			memcpy(Make_Temp.yt1,p_y.yt1,sizeof(p_y.yt1));			   
			memcpy(Make_Temp.yt2,p_y.yt2,sizeof(p_y.yt2));			   


			Make_Temp.th1_sum=p_y.th1_sum;
			Make_Temp.th2_sum=p_y.th2_sum;
			Make_Temp.sj1_sum=p_y.sj1_sum;
			Make_Temp.sj2_sum=p_y.sj2_sum;
			Make_Temp.yz1_sum=p_y.yz1_sum;
			Make_Temp.yz2_sum=p_y.yz2_sum;
			Make_Temp.yt1_sum=p_y.yt1_sum;
			Make_Temp.yt2_sum=p_y.yt2_sum;

			Make_Temp.yz1_sum=p_y.yz1_sum;
			for(i=0;i<Make_Temp.th1_sum;i++)//����̼��1��������
	                	Make_Temp.th1[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.th2_sum;i++)//����̼��2��������
       	         	Make_Temp.th2[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yz1_sum;i++)//��������һ1��������
              	  	Make_Temp.yz1[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yz2_sum;i++)//��������һ1��������
              	  	Make_Temp.yz2[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yt1_sum;i++)//��������һ1��������
              	  	Make_Temp.yt1[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yt2_sum;i++)//��������һ1��������
              	  	Make_Temp.yt2[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.sj1_sum;i++)//��������һ1��������
              	  	Make_Temp.sj1[i].get_weight=60.25f;
			for(i=0;i<Make_Temp.sj2_sum;i++)//��������һ1��������
              	  	Make_Temp.sj2[i].get_weight=60.35f;
			
			
			Make_Temp.get_weight=100.0f; //С��ʵ������
			Make_Temp.huan_time=10.0f;  //����ʱ��
			Make_Temp.pai_tempro=100;  //����ʱ��
			Make_Temp.all_power=400.0f;  //����������
			MessageBox(hdc,"дͳ������","Test ",MB_OK);
		//	write_che(Make_Temp);	//дһ��ͳ������
			Sleep(6000);
        }
   Sleep(1000);
   }while(1);
return FALSE;
}
/**********************************************************
���Զ��豸�����߳�   
**********************************************************/
LRESULT WINAPI Read_Comm_Thread (int Y)
{
char temp[280];
struct  send_table  x;
int i=1;			
 do{
	HDC  hdc;
    //   if(read_send_table(2,1,&x)==FALSE)			   
	  // sprintf(temp,"��send_table ʧ�� ");						
	//else
	  // sprintf(temp,"�豸��=%s ��Сֵ=%f ���ֵ=%f ���AD=%d,��СAD=%d ��=%d",
		//	x.name,x.min,x.max,x.AD_max,x.AD_min,x.port);
	hdc=GetDC(hwnd);
       i=strlen(temp);
       TextOut(hdc,40,120,temp,i);				
	ReleaseDC(hwnd,hdc);
	Sleep(2000);
 }while(1);
 return FALSE;
}


/*
 * FirstInstance - ΪӦ�ó���ǼǴ���Class �������ʼ��                 
 */
static BOOL FirstInstance( HANDLE this_inst )
{
    WNDCLASS    wc;
    BOOL        rc;
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = (LPVOID) WindowProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = sizeof( DWORD );
    wc.hInstance = this_inst;
    wc.hIcon = LoadIcon( this_inst, "GenericIcon" );
    wc.hCursor = LoadCursor( NULL, IDC_ARROW );
    wc.hbrBackground = GetStockObject( WHITE_BRUSH );
    wc.lpszMenuName = "GenericMenu";
    wc.lpszClassName = GenericClass;
    rc = RegisterClass( &wc );
    return( rc );

} /* FirstInstance */

/*
 * AnyInstance - ����Windows ����
 */
static BOOL AnyInstance( HANDLE this_inst, int cmdshow, LPSTR cmdline )
{

    extra_data  *edata_ptr;    
    hwnd = CreateWindow(
        GenericClass,           /* class */
        "	���Գ���",   /* caption */
        WS_OVERLAPPEDWINDOW,    /* style */
        CW_USEDEFAULT,          /* init. x pos */
        CW_USEDEFAULT,          /* init. y pos */
        CW_USEDEFAULT,          /* init. x size */
        CW_USEDEFAULT,          /* init. y size */
        NULL,                   /* parent window */
        NULL,                   /* menu handle */
        this_inst,              /* program handle */
        NULL                    /* create parms */
        );
                    
    if( !hwnd ) return( FALSE );

    edata_ptr = malloc( sizeof( extra_data ) );
    if( edata_ptr == NULL ) return( FALSE );
    edata_ptr->cmdline = cmdline;
    SetWindowLong( hwnd, EXTRA_DATA_OFFSET, (DWORD) edata_ptr );

    /*
     * ��ʾ����
     */
    ShowWindow( hwnd, cmdshow );
    UpdateWindow( hwnd );
    
    return( TRUE );
                        
} /* AnyInstance */

/*
 * AboutDlgProc - �Ի���  
 */


BOOL _EXPORT FAR PASCAL AboutDlgProc( HWND hwnd, unsigned msg,UINT wparam, LONG lparam )
{
    lparam = lparam;                    /* turn off warning */
    switch( msg ) {
    case WM_INITDIALOG:
        return( TRUE );

    case WM_COMMAND:
        if( LOWORD( wparam ) == IDOK ) {
            EndDialog( hwnd, TRUE );
            return( TRUE );
        }
        break;
    }
    return( FALSE );

} /* AboutDlgProc */

/*
 * WindowProc - Ӧ�ó������Ϣ����
 */
LONG _EXPORT FAR PASCAL WindowProc( HWND hwnd, unsigned msg,
                                     UINT wparam, LONG lparam )
{
    FARPROC     proc;
    extra_data  *edata_ptr;
    char        buff[128];
    int i,x;

    DWORD    dwThreadID;


    switch( msg ) {
	    case WM_CREATE:
	       init_elec_input(&elec_input);
        	init_elec_output(&elec_output);
           break;
    case WM_COMMAND:
        switch( LOWORD( wparam ) ) 
		{
		case MENU_KILL:	//ɱ�̵߳�����
		for(i=0;i<total_thread;i++)		 
		{			
			TerminateThread(chdthd[i],1);
		}
		total_thread=0;
		Z_DLL_Close_All();
	       i=sio_close(TEST_PORT);			
		 break;         
        case MENU_ABOUT:
            proc = MakeProcInstance( AboutDlgProc, MyInstance );
            DialogBox( MyInstance,"AboutBox", hwnd, proc );
            FreeProcInstance( proc );
            break;

        case MENU_CREATE://�����̵߳�����
		if (total_thread>0){
		MessageBox(0,"�Ѿ��������߳�","",0);
		MessageBeep(0);	
		break;
	}		
	  hThread = CreateThread(NULL, 0,                            
          (LPTHREAD_START_ROUTINE) Read_Elec_Thread,
                            (LPVOID)i,0, &dwThreadID);       
	if  (hThread==0)
			MessageBox(0,"create Read_Elec_Thread thread faile","",0);
	else
	 	 chdthd[total_thread++]=hThread		;
	  hThread = CreateThread(NULL, 0,                            
                            (LPTHREAD_START_ROUTINE) Read_Comm_Thread,
                            (LPVOID)i,0, &dwThreadID);       
		if  (hThread==0)
			MessageBox(0,"create scond thread faile","",0);
	else
	 	 chdthd[total_thread++]=hThread		;
          if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) Write_Elec_Thread,
                            (LPVOID)i,0, &dwThreadID)))
       		MessageBox(0,"create Write_Elec_Thread thread faile","",0);
	   else
	        chdthd[total_thread++]=hThread;

         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) Read_Run_Thread,
                            (LPVOID)i,0, &dwThreadID)))
       		MessageBox(0,"create Write_Elec_Thread thread faile","",0);
	  else
            chdthd[total_thread++]=hThread;
        break;

	case MENU_COM_READ:
	    hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) read_Thread,
                            (LPVOID)i,
                             0, &dwThreadID);
		chdthd[total_thread++]=hThread	;	
		 break;
	case MENU_COM_WRITE:
	    hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) write_Thread,
                            (LPVOID)i,
                             0, &dwThreadID);
		chdthd[total_thread++]=hThread	;	
		 break;
	case Menu_Read_Xiao:
		{
			struct  Produce_type py;
			struct  run_table  run_table_do;
			int run_flag;
			HDC hdc;
			char temp[80];
	       read_run(&run_table_do, &run_flag);
		   if (run_flag!=STOP_DOING)
		   {
				read_produce(&py);
		   }
		   else
			   memset(&(py.base),0,sizeof(py.base));
			hdc=GetDC(hwnd);
	       sprintf(temp,"����ʱ��=%d С�Ϲ���=%f",
			   py.base.temp1,py.base.temp2);
           TextOut(hdc,60,120,temp,strlen(temp));					  
		   ReleaseDC(hwnd,hdc);
		}
		break;
	case Menu_Read_Fen_Ban:
		{
			HDC hdc;
			char temp[200];
			struct  ban_table ban[45];
				read_fen_ban(ban);
				i=0;
				ban[i]._number[0]=0;
				hdc=GetDC(hwnd);

				for(i=0;i<40&ban[i]._number[0]!='\0';i++)
				{
				sprintf(temp,"��������Ϊ: ����%s,����%s,��%d,��%d,��%d,�ܳ�%d",ban[i]._number,
					ban[i].name,ban[i].ban,ban[i]._che,ban[i].mathine,ban[0].total_che);
				TextOut(hdc,20,30+i*20,temp,strlen(temp));					  
				}
			   ReleaseDC(hwnd,hdc);
		}
		break;
	}//end case
	break;
    case WM_DESTROY:				
         x=sio_close(TEST_PORT);
		 Z_DLL_Close_All();
        PostQuitMessage( 0 );
        break;
    default:
        return( DefWindowProc( hwnd, msg, wparam, lparam ) );
    }
    return( 0L );

} /* WindowProc */


/*****************************************************
	炼胶生产系统:生产测试程序
	设计者      :赖健敏,朱国魂
	日期	     :1998.1.12
	编译器要求  :Vitual C++ 或Bland C++
	环境要求    :炼胶软件已安装
	测试内容    :多用户卡通讯 ,多屏卡,多线程,配方传递 ,生产统计子程序
		   	测试读设备参数.本程序的测试数据已通过.
*****************************************************/
#include <windows.h>
#include <stdio.h>
#include <malloc.h>
#include "Resrc1.h"
#include "genwin.h"
#include "dll.h"
#include "api232-c.h"

#define TEST_LEN    (1024 * 10)		//测试Buffer 长度
#define TEST_PORT   2			//测试 口 2
int  WINAPI can_read(BSTR * pei_number,BSTR *name,short ban,char total_che,short flag);
HANDLE   hThread;				//主线程   
HANDLE    chdthd[10];			//测试子线程   最多10个 
HWND      hwnd;				//主线程 hwnd
int     total_thread=0;			//当前测试子线程个数
int     run_flag=STOP_DOING;      		//运行标志

/*****************************************************
	Windows 常规申明  可参考一般Windows程序书
*****************************************************/
HANDLE          MyInstance;
static char     GenericClass[32]=" 测试程序";
static BOOL FirstInstance( HANDLE );
static BOOL AnyInstance( HANDLE, int, LPSTR );
long  FAR PASCAL WindowProc( HWND, unsigned, UINT, LONG );
/*****************************************************
	Windows 主函数  可参考一般Windows程序书
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
	写通讯口测试线程 
*****************************************************/
LRESULT WINAPI write_Thread (int x)
{
	     char  temp[TEST_LEN+1];
	     int i=0;
	     HDC   hdc;
	       if (sio_open(TEST_PORT)<0)	//测试函数 sio_open		
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
			sprintf(temp,"%8s","睡眠 1s");
			TextOut(hdc,10,10,temp,8);
			ReleaseDC(hwnd,hdc);
			Sleep(1000);		//睡眠1000ms  即1S
			i=1;
		}
	}
 return FALSE;
}
/*****************************************************
读通讯口测试线程   从PLC读出  由赖健敏,朱国魂完成
*****************************************************/
LRESULT WINAPI read_Thread (int x)
{
    char  temp[40],temp1[40],s1[40],s2[40];
    int i=0,j,k;
    unsigned c1;
    float temp_float,f1;
    static int static_d=0;	
    HDC  hdc;

      sio_close(2);	//为安全 先关闭 2口
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
		//测试 写秤函数
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
测试读电信号线程   从控制程序读  
*****************************************************/
struct  elect_send_type     elec_input;  //  测试电信号  输入信号声明
struct  elect_send_type     elec_output;  // 测试电信号  输入信号声明
LRESULT WINAPI Read_Elec_Thread (int x)
{
	char temp[280];
		HDC  hdc;
	 do{
         if (run_flag!=STOP_DOING)
         {			   		   
		int x;
		//控制程序初始化 电信号
//  	       init_elec_input(&elec_input);
//	       init_elec_output(&elec_output);
		
		   hdc=GetDC(hwnd);
		   sprintf(temp,"输入信号总长度=%d 输出信号总长度=%d",
			elec_input.length,elec_output.length);
		   x=strlen(temp);
	          TextOut(hdc,60,120,temp,x);					  
		   ReleaseDC(hwnd,hdc);
		   write_multi_screen(120,60,temp);	//写多屏卡
        }
   	Sleep(1000);
 }while(1);
return FALSE;
}

/*****************************************************
测试风送系统线程   从控制程序读  
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
	    sprintf(temp,"图库代号为:%s",grap_name);
		x=strlen(temp);
		TextOut(hdc,20,10,temp,x);					  		
		if (strcmp(grap_name,"main_tk")!=0)
		{

			if (run_flag==STOP_DOING)
			{			   		   
		//		read_fen_ban(ban);
		/*		for(i=0;i<40&ban[i]._number[0]!='\0';i++)
				{
				sprintf(temp,"生产命令为: 材料%s,名称%s,班%d,车%d,机%d,总车%d",ban[i]._number,
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
		   sprintf(temp,"第2 机组 风送设备"
			);
			x=strlen(temp);
	        TextOut(hdc,60,120,temp,x);					  

			for(i=0;i<fx[0].Fsb_jzh;i++)
			{
			    sprintf(temp,"斗号=%d 品种号=%d,单罐=%f,工作压力=%f,中压=%f,低压=%f,高压=%f,输送=%d,清扫=%d,停用=%d,比例=%d,备用=%d",
				fx[i].Fsb_dh,fx[i].Fsb_pzh,fx[i].Fsb_dwight,fx[i].Fsb_gzyl,fx[i].Fsb_zysd,fx[i].Fsb_dysd,fx[i].Fsb_dsgy,fx[i].Fsb_mtime,
				fx[i].Fsb_qtime,fx[i].Fsb_ty,fx[i].Fsb_blxs,fx[i].Fsb_byl);
			    x=strlen(temp);
				TextOut(hdc,20,140+i*20,temp,x);					  					
			}


//		read_Fsml_table(fs,&flag,&number);
		hdc=GetDC(hwnd);
		   sprintf(temp,"风送系统个数=%d 运行命令=%d",
			number,flag);
	    TextOut(hdc,20,220,temp,strlen(temp));					  

       if(flag!=0)
	   {
		   int send1,time1;
//		   write_Fsml_table(fs,2,number);
//		   write_chans(1,2,3,4);
//		   write_runfch(data+2,data,2*data,data*4,data+3,"测试");
		   send1=data+2;
			time1=data+3;
		   sprintf(temp,"输送=%4d,1#=%4.6f 2#=%4.6f , 压力=%4.6f,时间=%4d",
			send1,data,2*data,data*4,time1);
		    TextOut(hdc,20,240,temp,strlen(temp));					  
		   data+=2.4;
	   }
	   if (data>100)
		   data=0.5;
	   

		read_fen_ban(ban);
		for(i=0;i<40&ban[i]._number[0]!='\0';i++)
		{
		    sprintf(temp,"生产命令为: 材料%s,名称%s,班%d,车%d,机%d,总车%d",ban[i]._number,
				ban[i].name,ban[i].ban,ban[i]._che,ban[i].mathine,ban[0].total_che);
			TextOut(hdc,20,30+i*20,temp,strlen(temp));					  

		}

	   for(i=0;i<number;i++)
	   {
		   sprintf(temp,"序号=%d 机组=%d,斗号=%d,品种号=%d,材料=%s,名称=%s,输送=%d,连送=%d,备用=%d",
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
测试生产统计数据线程   须先运行生产命令 ,否则不会有数据 显示和产生
**********************************************************/
struct make_base_type Make_Temp;	//  测试每段 数据声明
struct make_lian_liao Duan_Data;	//  测试每车 数据声明
LRESULT WINAPI Read_Run_Thread (int x)
{
struct  run_table run_table_do;
struct  Produce_type p_y;
int i; char temp[280];
    do{
       static int doing_che=0;       
       char   *statu[6]={"停止", "运行","模拟"};
  //     read_run(&run_table_do, &run_flag);
       if (run_flag!=STOP_DOING)
       {			   	
	   	HDC  hdc;
	       hdc=GetDC(hwnd);
	       sprintf(temp,"总车数=%d,机号=%d,配方编号=%s,配方名称=%s",run_table_do.total_che,
		run_table_do.mathine,run_table_do._number,run_table_do.name);
	       run_table_do.now_che++;//当前车号++				
  	       if(run_table_do.now_che>run_table_do.total_che)//判断结束本配方否
              {
                  run_table_do.now_che=0;			//   
    //              write_run(&run_table_do,0);		//运行=0  表示结束	
	  //	    write_power(4,50,60,-1);  //write 温度,能量
		    sprintf(temp,"运行结束");					
		    x=strlen(temp);  TextOut(hdc,60,40,temp,x);				
		    return FALSE;
                }else
                    write_run(&run_table_do,-1);	//运行=-1  表示继续	
	        x=strlen(temp);
               TextOut(hdc,60,40,temp,x);				
	        write_multi_screen(160,60,temp);
               read_produce(&p_y );
		 sprintf(temp,"碳黑投1数=%d,油料1投1=%d,油料2投1=%d 胶料数=%d 总段数=%d",
		           p_y.th1_sum,p_y.yz1_sum,p_y.yt1_sum,p_y.sj1_sum,p_y.pei_max);
		 x=strlen(temp); TextOut(hdc,60,260,temp,x);				
	        sprintf(temp,"name=%s,油1重量=%f,允许公差=%f,fast_do=%f,drop_do=%f",p_y.yz1[0].name,p_y.yz1[0].weight,
			p_y.yz1[0].gon_cha,p_y.yz1[0].fast_do,p_y.yz1[0].drop_do);
		 x=strlen(temp);   TextOut(hdc,60,180,temp,x);				
	        sprintf(temp,"name=%s,油2重量=%f,允许公差=%f,fast_do=%f,drop_do=%f",p_y.yz2[0].name,p_y.yz2[0].weight,
			p_y.yz2[0].gon_cha,p_y.yz2[0].fast_do,p_y.yz2[0].drop_do);
		  x=strlen(temp);TextOut(hdc,60,200,temp,x);				
	         sprintf(temp,"胶料1=%2d,胶料2=%2d,碳黑1=%2d,碳黑2=%2d,油1=%2d,油2=%2d",
			p_y.sj1_sum,p_y.sj2_sum,p_y.th1_sum,p_y.th2_sum,p_y.yz1_sum,p_y.yz2_sum);
	         x=strlen(temp); TextOut(hdc,60,240,temp,x);				
			 if (p_y.yt1_sum>0)
			 {
				sprintf(temp,"name=%s,油21重量=%f,允许公差=%f,fast_do=%f,drop_do=%f",p_y.yt1[0].name,p_y.yt1[0].weight,
				p_y.yt1[0].gon_cha,p_y.yt1[0].fast_do,p_y.yt1[0].drop_do);
				x=strlen(temp);   TextOut(hdc,60,260,temp,x);				

			 }
			 if (p_y.yt2_sum>0)
			 {
		        sprintf(temp,"name=%s,油22重量=%f,允许公差=%f,fast_do=%f,drop_do=%f",p_y.yt2[0].name,p_y.yt2[0].weight,
				p_y.yt2[0].gon_cha,p_y.yt2[0].fast_do,p_y.yt2[0].drop_do);
				x=strlen(temp);TextOut(hdc,60,280,temp,x);				
			 }

		   ReleaseDC(hwnd,hdc);
		//   write_multi_screen(200,60,temp);
	
		   //测试一车开始
	/*先测试一车设置的所有段数据*/
	for(i=0;i<p_y.pei_max;i++)
	{	
				int j;
		//		for(j=0;j<30;j++)
		///		write_power(i+1,j+30,(j+30)%40,1);  //write 温度,能量

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
			MessageBox(hdc,"写段数据","Test ",MB_OK);
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
			for(i=0;i<Make_Temp.th1_sum;i++)//测试碳黑1生产数据
	                	Make_Temp.th1[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.th2_sum;i++)//测试碳黑2生产数据
       	         	Make_Temp.th2[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yz1_sum;i++)//测试油料一1生产数据
              	  	Make_Temp.yz1[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yz2_sum;i++)//测试油料一1生产数据
              	  	Make_Temp.yz2[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yt1_sum;i++)//测试油料一1生产数据
              	  	Make_Temp.yt1[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.yt2_sum;i++)//测试油料一1生产数据
              	  	Make_Temp.yt2[i].get_weight=(60.5f+i*2)/3;
			for(i=0;i<Make_Temp.sj1_sum;i++)//测试油料一1生产数据
              	  	Make_Temp.sj1[i].get_weight=60.25f;
			for(i=0;i<Make_Temp.sj2_sum;i++)//测试油料一1生产数据
              	  	Make_Temp.sj2[i].get_weight=60.35f;
			
			
			Make_Temp.get_weight=100.0f; //小料实际重量
			Make_Temp.huan_time=10.0f;  //混炼时间
			Make_Temp.pai_tempro=100;  //混炼时间
			Make_Temp.all_power=400.0f;  //消耗总能量
			MessageBox(hdc,"写统计数据","Test ",MB_OK);
		//	write_che(Make_Temp);	//写一车统计数据
			Sleep(6000);
        }
   Sleep(1000);
   }while(1);
return FALSE;
}
/**********************************************************
测试读设备参数线程   
**********************************************************/
LRESULT WINAPI Read_Comm_Thread (int Y)
{
char temp[280];
struct  send_table  x;
int i=1;			
 do{
	HDC  hdc;
    //   if(read_send_table(2,1,&x)==FALSE)			   
	  // sprintf(temp,"读send_table 失败 ");						
	//else
	  // sprintf(temp,"设备名=%s 最小值=%f 最大值=%f 最大AD=%d,最小AD=%d 口=%d",
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
 * FirstInstance - 为应用程序登记窗口Class 和其余初始化                 
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
 * AnyInstance - 产生Windows 窗口
 */
static BOOL AnyInstance( HANDLE this_inst, int cmdshow, LPSTR cmdline )
{

    extra_data  *edata_ptr;    
    hwnd = CreateWindow(
        GenericClass,           /* class */
        "	测试程序",   /* caption */
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
     * 显示窗口
     */
    ShowWindow( hwnd, cmdshow );
    UpdateWindow( hwnd );
    
    return( TRUE );
                        
} /* AnyInstance */

/*
 * AboutDlgProc - 对话框  
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
 * WindowProc - 应用程序的消息处理
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
		case MENU_KILL:	//杀线程的命令
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

        case MENU_CREATE://产生线程的命令
		if (total_thread>0){
		MessageBox(0,"已经产生了线程","",0);
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
	       sprintf(temp,"返料时间=%d 小料公差=%f",
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
				sprintf(temp,"生产命令为: 材料%s,名称%s,班%d,车%d,机%d,总车%d",ban[i]._number,
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


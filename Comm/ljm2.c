#include <windows.h>
#include <stdio.h>
#include <math.h>
#include <time.h>
#include <malloc.h>
#include "genwin.h"
#include "dll.h"
#include "api232-c.h"
#define TEST_LEN    (1024 * 10)
#define TEST_PORT   2
	struct  run_table run_table_do;
	struct  Produce_type p_y;
	struct  elect_send_type     elec_input;  //elec
	struct  elect_send_type     elec_output;  //elec
	struct  make_lian_liao Duan_Data;
	struct  make_base_type Make_Temp;

int h16_bcd(int i);
int bcd_h16(int i);
char *h16_b02(int i);
char *ii_bin(unsigned int i,char *p);
char *i_bin(unsigned int i,char *p);
char *h16_16a(unsigned int i,char *p);
char *head_d(char *p,char *p1,unsigned int m,unsigned int i);
char *head_h(char *p,char *p1,unsigned int m,unsigned int i);
unsigned asc_10i(char *p);
unsigned asc_16i(char *p);
char *ad_d(char *p);
char *head_t(char *p,char *p1);
void bcom1(char *p);
void bcom2();
void bcom3();
void bcom4();
void read_D();
void read_X();
void read_Y();
void read_M1();
void read_M2();
void read_D_end();
void write_Y();
void write_p_f();

char d_1[49],d_2[53],p_d[14][5];
char d_3[6];
int  d_d[21],cs_flag=0;
int	 data_d[70],data_m1[128];//data_m2[66];
float d_f;
double tim1,tim2;
time_t tim_1,tim_2,tim_3;


HANDLE   hThread;
HANDLE    chdthd[10];
LRESULT CALLBACK  ThreadWndProc(HWND hwnd,UINT msg,
                                     UINT wparam, LONG lparam );
HWND        hwnd;//main  hwnd

int     total_thread=0;
int     run_flag=STOP_DOING;      //run_flag


HANDLE          MyInstance;
static char     GenericClass[32]=" 测试程序";

static BOOL FirstInstance( HANDLE );
static BOOL AnyInstance( HANDLE, int, LPSTR );

long  FAR PASCAL WindowProc( HWND, unsigned, UINT, LONG );

/*
 * WinMain - initialization, message loop
 */
HANDLE hsem;
int PASCAL WinMain( HANDLE this_inst, HANDLE prev_inst, LPSTR cmdline,
                    int cmdshow )
{
    MSG         msg;
	int			i;

    MyInstance = this_inst;
	hsem=CreateSemaphore(NULL,1,1,"ZGH");
	if (GetLastError()==ERROR_ALREADY_EXISTS)
	{
		MessageBox(0,"已经运行 ","错误",0);		
		CloseHandle(hsem);
		return FALSE;
	}
	

#ifdef __WINDOWS_386__
    sprintf( GenericClass,"GenericClass%d", this_inst );
    prev_inst = 0;
#endif
	
    if( !prev_inst ) {
        if( !FirstInstance( this_inst ) ) return( FALSE );
    }
    if( !AnyInstance( this_inst, cmdshow, cmdline ) ) return( FALSE );

    while( GetMessage( &msg, NULL, NULL, NULL ) ) {

        TranslateMessage( &msg );
        DispatchMessage( &msg );

    }
		for(i=0;i<total_thread;i++)		 
		{			
			TerminateThread(chdthd[i],1);
		}
		total_thread=0;
		  i=sio_close(TEST_PORT);
//		  Sio_Open_All( );

     return FALSE;

    return( msg.wParam );

} /* WinMain */


/*LRESULT WINAPI write_Thread (int x)
{
     char  temp[TEST_LEN+1];
     int i;
	 HDC   hdc;
		i=0;
		  x=sio_open(TEST_PORT);
		  if (x<0)
		  {
			   //MessageBox(hwnd,"open  error","open",0);
//				return  0;
		  }
			wsprintf(temp, "call sio_ioctl return value = %d", sio_ioctl(TEST_PORT,
				B9600, BIT_8 | STOP_1 | P_NONE));
			   MessageBox(hwnd,temp,"ioctol ",0);
             while (TRUE)
				{
					i++;
                    wsprintf(temp,"%10d",i);
                    x=sio_write(TEST_PORT, temp, TEST_LEN);
					if(x<0) 
					{
							MessageBox(hwnd,"write  error","open",0);
							return ;
					}
						
					hdc=GetDC(hwnd);
                    wsprintf(temp,"write %10d return %d ",i,x);
                    TextOut(hdc,10,10,temp,40);
					ReleaseDC(hwnd,hdc);

					if(i>4000)
					{											
						hdc=GetDC(hwnd);
						sprintf(temp,"%8s","睡眠 1s");
						TextOut(hdc,10,10,temp,8);
						ReleaseDC(hwnd,hdc);
						Sleep(1000);
						i=1;
					}
				}
				;

                return FALSE;
}

LRESULT WINAPI read_Thread (int x)
{
    char  temp[40],temp1[40];
    int i=0,j,k;
	float f1;
	char s1[40],s2[40];
	unsigned c1;
	float temp_float;
	

	HDC  hdc;

	return FALSE;

}*/



LRESULT WINAPI OneThread (int x)
{
	int i,k,now_dh=0,R_io;//now_che;
	char temp[80];
	float nl;
	int  run_flag;
	HDC  hdc;


		hdc=GetDC(hwnd);
	while(1)
	{

	do
	 {
		init_elec_input(&elec_input);
		init_elec_output(&elec_output);
		if(elec_output.flag != 2)
		{
			read_X();
			read_Y();
			read_D();
		}
		else
			write_Y();

		read_run(&run_table_do, &run_flag);
//			sprintf(temp,"%4x",elec_output.flag);//run_flag
//			TextOut(hdc,50,40,temp,4);
//			sprintf(temp,"%4x",cs_flag);
//			TextOut(hdc,50,80,temp,4);
	  Sleep(880);
     }while(run_flag==STOP_DOING);
	 write_p_f();
	 write_run(&run_table_do,-1);
	 run_table_do.now_che=1;//now_che=1;
	 sprintf(temp,"运行开始  ");
	 x=strlen(temp);
	 TextOut(hdc,20,50,temp,x);
	 write_run(&run_table_do,-1);
	 while(run_table_do.now_che <= run_table_do.total_che)
	 {
		read_run(&run_table_do, &run_flag);
		Sleep(110);
//		sprintf(temp,"run=%3d",run_flag);
//		TextOut(hdc,20,20,temp,20);
		if(run_flag==3)
		{
			write_p_f(); 
			write_run(&run_table_do,4);
		}
		 now_dh=1;nl=1.0;
		 time(&tim_3);
		 run_table_do.all_che_time=1;
		 while(now_dh <= p_y.pei_max)
		 {
			time(&tim_1);
			run_table_do.duan_time=1;
			do{
			//	if(R_io==0)
			//	{read_X();R_io=1;}
			//	else
			//	{read_Y();R_io=1;}
			//	read_D();
				time(&tim_2);
				tim1=difftime(tim_2,tim_1);
			    tim2=difftime(tim_2,tim_3);
				nl=nl+2.0;
				k=(int)nl+5;
				write_power(now_dh,nl,k,1);	//写温度、能量进行图表统计				 
				sprintf(temp,"%3d -- %3d 总时间=%3d",
						(int)tim1,p_y.huan[i].lian_time,(int)tim2);
				TextOut(hdc,120,50,temp,strlen(temp));
				run_table_do.duan_time=(int)tim1;
				run_table_do.all_che_time=(int)tim2;
				write_run(&run_table_do,-1);
				Sleep(880);
			 }while((int)tim1<p_y.huan[i].lian_time);
			nl=0.0;
	//		run_table_do.duan_time=-1;
	//		write_run(&run_table_do,-1);
		 	Duan_Data.duan_hao=now_dh;
			Duan_Data.get_hun_time=(int)tim1;
			Duan_Data.get_you_time=p_y.huan[i].dou_you_time;
			Duan_Data.get_you2_time=p_y.huan[i].dou_you2_time;
			Duan_Data.get_jiao_time=p_y.huan[i].dou_jiao_time;
			Duan_Data.get_tan_time=p_y.huan[i].dou_tan_time;
	//		Duan_Data.get_xiao_time=p_y.huan[i].dou_xiao_time;
						
			Duan_Data.next_tempro=k;
			Duan_Data.next_power=nl;
							
	//		lstrcpy(Duan_Data.duan_begin_time,"11:12:12");
	//		lstrcpy(Duan_Data.duan_end_time,"11:22:12");

			write_make(Duan_Data);
			sprintf(temp,"%2d 段结束",now_dh);
			x=strlen(temp);
			TextOut(hdc,320,50,temp,x);
			now_dh++;
		 }
		time(&tim_2);
		tim2=difftime(tim_2,tim_3);
		run_table_do.all_che_time=(int)tim2;
		write_power(now_dh,nl,k,-1);	// 一车结束，结束能量、温度统计
		Make_Temp.pei_max=i;
		memcpy(Make_Temp.sj1,p_y.sj1,sizeof(p_y.sj1));			   
		memcpy(Make_Temp.sj2,p_y.sj2,sizeof(p_y.sj2));			   
		memcpy(Make_Temp.th1,p_y.th1,sizeof(p_y.th1));			   
		memcpy(Make_Temp.th2,p_y.th2,sizeof(p_y.th2));			   
		memcpy(Make_Temp.yz1,p_y.yz1,sizeof(p_y.yz1));			   
		Make_Temp.sj1_sum=p_y.sj1_sum;
		Make_Temp.sj2_sum=p_y.sj2_sum;
		Make_Temp.th1_sum=p_y.th1_sum;
		Make_Temp.th2_sum=p_y.th2_sum;
		Make_Temp.yz1_sum=p_y.yz1_sum;
		for(i=0;i<Make_Temp.sj1_sum;i++)//测试胶料1生产数据
				Make_Temp.sj1[i].get_weight=p_y.sj1[i].weight+2;
		for(i=0;i<Make_Temp.sj2_sum;i++)//测试胶料2生产数据
				Make_Temp.sj2[i].get_weight=p_y.sj2[i].weight+2;

		for(i=0;i<Make_Temp.th1_sum;i++)//测试碳黑1生产数据
				Make_Temp.th1[i].get_weight=p_y.th1[i].weight+2;
		for(i=0;i<Make_Temp.th2_sum;i++)//测试碳黑2生产数据
				Make_Temp.th2[i].get_weight=p_y.th2[i].weight+2;

		for(i=0;i<Make_Temp.yz1_sum;i++)//测试油料一1生产数据
				Make_Temp.yz1[i].get_weight=p_y.yz1[i].weight+2;
                
		Make_Temp.get_weight=100.0f; //小料实际重量
		Make_Temp.huan_time=(int)tim2;  //混炼总时间
		Make_Temp.pai_tempro=102;  //排料温度
		Make_Temp.get_pai_time=103;//排料时间
		Make_Temp.all_power=104.0f;  //消耗总能量
		write_che(Make_Temp);	//写一车统计数据
		run_table_do.all_che_time=1;
	//	write_run(&run_table_do,-1);
		sprintf(temp,"%2d 车结束",run_table_do.now_che);
		x=strlen(temp);
		TextOut(hdc,420,50,temp,x);
		run_table_do.now_che++;
		write_run(&run_table_do,-1);		
		Sleep(880);

	 }
	 write_run(&run_table_do,0);
	 cs_flag=0;
	 sprintf(temp,"运行结束");
	 x=strlen(temp);
	 TextOut(hdc,30,50,temp,x);
	 Sleep(1100);
//	if(data_m1[90]==1) {};
	{
/**************************
	 add by zgh 4.1
	 
**************************/
	/*	   if (run_flag==3)
		   {
			   sprintf(temp,"在线修改%d车结束",run_table_do.now_che);
			   x=strlen(temp);
			   TextOut(hdc,30,80,temp,x);
			   if(run_table_do.now_che>run_table_do.total_che)
			   {
				   sprintf(temp,"在线修改%d车非法",run_table_do.total_che);
				   x=strlen(temp);
				   TextOut(hdc,30,80,temp,x);
			   }	
			   else
			   		read_produce(&p_y );

				write_run(&run_table_do,4);		
		   }*/
/**************************
	 up add by zgh 4.1
	 
**************************/
		//	now_che++;
	 }
        ReleaseDC(hwnd,hdc);
		sio_flush(2,1);
	}
	return FALSE;
}

LRESULT WINAPI TwoThread (int x)
{
	int i=0,tj,yj,jl;
//	char *s1,*s2;
	char zoj[20];
	char temp[40];

	float f1,f2,f3,f4;
	HDC  hdc;
    while (TRUE)
		{
        hdc=GetDC(hwnd);
//		do
//		{
//			Sleep(110);
//		}while(cs_flag==0);
//		run_table_do.all_che_time=1;
//		run_table_do.duan_time=1;
//		write_run(&run_table_do,-1);
		f1=9.8;
		for(i=0;i<10;i++)
		{
			f2=i*8.2;
			f3=(float)(data_d[5]/100.0); //i*2.9;
			f4=i*12.6;
	//		sprintf(temp,"data_d=%4d f3=%f",data_d[5],f3);
	//		TextOut(hdc,10,10,temp,20);
			if(data_m1[21]==1 && data_m1[22]!=1) tj=2;
			if(data_m1[21]!=1 && data_m1[22]==1) tj=3;
			if(data_m1[26]==1 && data_m1[22]==1) tj=4;
			yj=2;
			jl=2;
			write_ban(1,1,&f1,&f2,tj);
			write_ban(1,2,&f1,&f3,yj);
			write_ban(1,4,&f1,&f4,jl);
/*                elec_output.data[i]=1;
                  write_elec_output(&elec_output);
*/
//		write_power(i,i*1.3,i*5,1);
//		run_table_do.now_che=i;
//		write_run(&run_table_do,-1);
                write_error(1,"DFASDF");
                write_error(3000,"测试错误型号");
		Sleep(1100);
		}
//		run_table_do.duan_time=0;
//		run_table_do.all_che_time=0;
//		write_run(&run_table_do,-1);
        ReleaseDC(hwnd,hdc);
        Sleep(550);
	}
 /*   char  temp[40],temp1[40];
    int i=0,j,k;
	float f1;
	char s1[40],s2[40];
	unsigned c1;

         x=sio_close(2);
          Sleep(200);
          x=sio_open(2);
		  if (x<0)
		  {
			   MessageBox(hwnd,"open  error","open 2",0);
               return  0;
		  }
            wsprintf(temp, "call sio_ioctl return value = %d", sio_ioctl(2,
				B2400, BIT_7 | STOP_2 | P_EVEN));

				for(i=0;i<8;i++)
                    temp1[i]=' ';
                 temp1[i]='\0';
              while (TRUE)
               {
                hdc=GetDC(hwnd);
  	  		    sio_flush(2,0);	
                Sleep(100);

				do
				{
                       i=sio_getch(2);
    				   Sleep(10);
                }while(i!=2);

                sio_read(2,s1,9);
				c1=s1[0]&0x007;
				for(i=0;i<7;i++)
				   s2[i]=s1[i+3];

				k=atoi(s2);
				switch(c1)
				{
					case 0x01: f1=k*10.0;break;
					case 0x03: f1=k/10.0;break;
					case 0x04: f1=k/100.0;break;
					case 0x05: f1=k/1000.0;break;
					case 0x06: f1=k/10000.0;break;
				};
                   if((s1[1]&0x02)!=0) f1=-f1;
                   sprintf(temp1,"%8.4f",f1);
                   TextOut(hdc,10,200,temp1,8);
   				   sio_flush(2,0);
                   Sleep(200);

               };
                ReleaseDC(hwnd,hdc);
                x=sio_close(2);
                  if (x<0)
                  {
                       MessageBox(hwnd,"close  error","open",0);
                        return  0;
                 }*/
        ReleaseDC(hwnd,hdc);
            return FALSE;

}

LRESULT WINAPI ThreeThread (int x)
{
//	int i,j,k,t_1=0,t_2=0,now_dh=0;//now_che;
//	char zoj[20];
//	char temp[40];
//	float jt;
//	char D_sx[5][6];
//	HDC  hdc;
 while (TRUE)
	{
 //    hdc=GetDC(hwnd);
		Sleep(3300);
 //    ReleaseDC(hwnd,hdc);
 }
    return FALSE;
}


LRESULT WINAPI FourThread (int x)
{
// int i,j,k;
 char temp[20];
// HDC hdc;
//	hdc=GetDC(hwnd);
	 do
	 {
	  Sleep(3300);
     }while(1);

 //  ReleaseDC(hwnd,hdc);
//   write_multi_screen(200,60,temp);
	
/*	 do
	 {
	  Sleep(1100);
     }while(cs_flag==1);
	int i=0,j,k,t_1=0,t_2=0;
	char *s1,*s2;
	char zoj[20];
	char temp[40];
	HDC  hdc;
	      x=sio_close(2);
          Sleep(300);
          x=sio_open(2);
		  if (x<0)
		  {
			   MessageBox(hwnd,"open error","open 2",0);
               return  0;
		  }
            wsprintf(temp, "call sio_ioctl return value = %d", sio_ioctl(2,
				B9600, BIT_7 | STOP_1 | P_ODD));
			    sio_flush(2,1);	
                Sleep(100);
    
              while (TRUE)
               {
                hdc=GetDC(hwnd);
	 			head_t(d_1,"QW");
				i=0;
				head_d(d_1,"D",10,5);
				for(i=t_1;i<t_1+5;i++)
				strcat(d_1,ii_bin(i,zoj));
				ad_d(d_1);
    //            TextOut(hdc,10,200,d_1,strlen(d_1));
				bcom1(d_1);
	//		    sio_flush(2,2);	
    //            Sleep(2);
	//			for(i=0;i<5;i++)
	//			{
				head_t(d_1,"QR");
				head_d(d_1,"D",10,5);
				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
	//			sio_write(2,d_1,strlen(d_1));
				bcom3();
				for(i=1;i<6;i++)
				{
				sprintf(temp,"%4d",d_d[i]);
                TextOut(hdc,400,(20+i*20),temp,4);
				}
				head_t(d_1,"QR");
				head_d(d_1,"D",1,5);

				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
				bcom4();
				sprintf(temp,"%6.2f",d_f);
                TextOut(hdc,400,250,temp,6);
				 

				head_t(d_1,"QW");
				head_d(d_1,"M",0,1);
				strcat(d_1,h16_16a(0xfae7,zoj));
				ad_d(d_1);
   //             TextOut(hdc,10,80,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"QR");
				head_d(d_1,"M",0,1);
				ad_d(d_1);
 //               TextOut(hdc,10,150,d_1,strlen(d_1));
				bcom1(d_1);
				bcom2();
				sprintf(temp,"%4x",d_d[1]);
                TextOut(hdc,300,40,temp,4);
				head_t(d_1,"JW");
				i=800;if(j == 1) j=0; else j=1;
				head_d(d_1,"M",i,16);
				for(i=0;i<16;i++)
				{if(j==0) j=1; else j=0;
				strcat(d_1,i_bin(j,zoj));}
				ad_d(d_1);
   //             TextOut(hdc,10,250,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"QR");
				head_d(d_1,"M",800,1);
				ad_d(d_1);
//                TextOut(hdc,10,200,d_1,strlen(d_1));
				bcom1(d_1);
				bcom2();
				k=d_d[1];
				sprintf(temp,"%4x",d_d[1]);
                TextOut(hdc,300,120,temp,4);


				head_t(d_1,"QW");
				head_h(d_1,"Y",0xe0,1);
				strcat(d_1,h16_16a(t_2,zoj));
				ad_d(d_1);
     //           TextOut(hdc,10,80,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"QR");
				head_h(d_1,"Y",0xe0,1);
				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
				bcom2();
				sprintf(temp,"%4x",d_d[1]);
                TextOut(hdc,300,80,temp,4);

				t_1=t_1+5;
				t_2=t_2+0x317;
				if(t_1>2000) t_1=0;
				if(t_2>0xfff0) t_2=0;
	//			sio_write(2,d_1,strlen(d_1));
			    sio_flush(2,1);	

                ReleaseDC(hwnd,hdc);
                Sleep(55);

               };*/
		//		sprintf(temp,"%4x",d_d[1]);
	/*	for(i=0;i<13;i++)
		{
			j=i*5+600;
			head_t(d_1,"QR");
			head_d(d_1,"D",j,5);
			ad_d(d_1);
			bcom1(d_1);
			bcom3();
			for(j=1;j<6;j++)
			{
			data_d[i*5+j]=d_d[j];
	//		sprintf(temp,"%4d",data_d[i*5+j]);
	//		TextOut(hdc,400,(20+j*20),temp,4);
	//		Sleep(550);
			}
		}
	*/
		//		head_t(d_1,"QW");
		//		head_d(d_1,"M",592,2);
		//		strcat(d_1,h16_16a(0xfae7,zoj));
		//		strcat(d_1,h16_16a(0x5f76,zoj));
		//		ad_d(d_1);
        //        TextOut(hdc,10,80,d_1,strlen(d_1));
    	//		bcom1(d_1);
	/*		head_t(d_1,"QR");
			head_d(d_1,"M",592,2);
			ad_d(d_1);
			bcom1(d_1);
			bcom2();
	//		sprintf(temp,"%4x",d_d[1]);
	//		TextOut(hdc,10,100,temp,4);
			for(j=0;j<16;j++)
			{	if((d_d[1]&(short)pow(2,j))>0)
				data_m2[j]=1;
				else
				data_m2[j]=0;
		//		sprintf(temp,"%4d",data_m2[j]);
		//		TextOut(hdc,100,(20+j*20),temp,4);
			}
			for(j=0;j<16;j++)
				if((d_d[2]&(unsigned short)pow(2,j))>0)
				data_m2[j+16]=1;
				else
				data_m2[j+16]=0;
		*/
			  Sleep(2200);
            return FALSE;
}

//test write_power   write_elec_input  
//                   write_elec_output
//	struct  elect_send_type     elec_input;  //elec
//	struct  elect_send_type     elec_output;  //elec


LRESULT WINAPI Read_Elec_Thread (int x)
{
		char temp[280];
	 do{
       if (run_flag==STOP_DOING)
       {			   	
		   	HDC  hdc;
			int x,i;
		    init_elec_input(&elec_input);
			init_elec_output(&elec_output);
		

			   hdc=GetDC(hwnd);
			   sprintf(temp,"输入信号总长度=%d 输出信号总长度=%d",
						elec_input.length,elec_output.length);

			   x=strlen(temp);
               TextOut(hdc,20,20,temp,x);				
				x=40;//j=1500;
			   for(i=0;i<elec_output.length;i++)
			   {
	//			   if(elec_input.note_name[i])
	//			   {
				   sprintf(temp,"节点名=%5d,数据=%2d",elec_output.note_name[i],elec_output.data[i]);
	               TextOut(hdc,320,x,temp,strlen(temp));
				   x=x+20;
	//			   }
			   }
			   ReleaseDC(hwnd,hdc);
			   write_multi_screen(120,60,temp);
        }
	  Sleep(2200);
	 }while(1);
        return FALSE;
}


LRESULT WINAPI Read_Run_Thread (int x)
{

	char temp[280];
	int i,j;
    do{
       static int doing_che=0;
       
 //      char   *statu[6]={"停止", "运行","模拟"};

 //      read_run(&run_table_do, &run_flag);
       if (run_flag!=STOP_DOING)
       {			   	
		   	HDC  hdc;
//			struct	elec_in_table elec_x[101];
			int x;
			   hdc=GetDC(hwnd);
/*		x=100;
		Get_Elec_Check_Table(elec_x,&x);
		for(i=0;i<x;i++)
		{
			int temp_x;
        sprintf(temp,"错误名=%2d,保留=%2s,启动接点=%2d,逻辑设定=%d,灯=%d,逻辑设=%2d,停检标志=%2d",
							elec_x[i].name,elec_x[i].str,elec_x[i].run,elec_x[i].run_set,
							elec_x[i].check,elec_x[i].check_set,elec_x[i].use_flag);
			   temp_x=strlen(temp);
               TextOut(hdc,10,80+i*20,temp,temp_x);	
		}*/
	//	read_run(&run_table_do, &run_flag);

			   sprintf(temp,"总车数=%2d,班号=%2d,机号=%2d,配方编号=%s,配方名称=%s,总段数=%2d,",
							run_table_do.total_che,run_table_do.ban,run_table_do.mathine,run_table_do._number,\
							run_table_do.name,run_table_do.all_duan);
			   x=strlen(temp);
               TextOut(hdc,10,80,temp,x);				
		//	   write_multi_screen(160,60,temp);
		   sprintf(temp,"碳投1=%2d,碳投2=%2d,油投1=%2d,油投2=%2d,胶1数=%2d,胶2数=%2d,总段数=%2d",
						p_y.th1_sum,p_y.th2_sum,p_y.yz1_sum,p_y.yz2_sum,p_y.sj1_sum,p_y.sj2_sum,p_y.pei_max);
				x=strlen(temp);
				TextOut(hdc,20,100,temp,x);j=100;
			   for(i=0;i<p_y.th1_sum;i++)
			   {
			   sprintf(temp,"斗号=%s,次数=%2d,碳投1重=%5.2f,快慢=%5.2f,点动=%5.2f,公差=%5.2f,控时=%3.2f,稳时=%2d",
						p_y.th1[i].dou,p_y.th1[i].add_time,p_y.th1[i].weight,p_y.th1[i].fast_do,\
						p_y.th1[i].drop_do,p_y.th1[i].gon_cha,p_y.th1[i].control_time,p_y.th1[i].stop_time);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }
			   for(i=0;i<p_y.th2_sum;i++)
			   {
			   sprintf(temp,"斗号=%s,次数=%2d,碳投2重=%5.2f,快慢=%5.2f,点动=%5.2f,公差=%5.2f,控时=%3.2f,稳时=%2d",
						p_y.th2[i].dou,p_y.th2[i].add_time,p_y.th2[i].weight,p_y.th2[i].fast_do,\
						p_y.th2[i].drop_do,p_y.th2[i].gon_cha,p_y.th2[i].control_time,p_y.th2[i].stop_time);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }
			   for(i=0;i<p_y.yz1_sum;i++)
			   {
			   sprintf(temp,"斗号=%s,次数=%2d,油投1重=%5.2f,快慢=%5.2f,点动=%5.2f,公差=%5.2f,控时=%3.2f,稳时=%2d",
						p_y.yz1[i].dou,p_y.yz1[i].add_time,p_y.yz1[i].weight,p_y.yz1[i].fast_do,\
						p_y.yz1[i].drop_do,p_y.yz1[i].gon_cha,p_y.yz1[i].control_time,p_y.yz1[i].stop_time);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }
			   for(i=0;i<p_y.sj1_sum;i++)
			   {
			   sprintf(temp,"次数=%2d,胶投1重=%5.2f,公差=%5.2f,胶状=%2d",
						p_y.sj1[i].add_time,p_y.sj1[i].weight,p_y.sj1[i].gon_cha,p_y.sj1[i].stop_time);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }
			   
			   for(i=0;i<p_y.pei_max;i++)
			   {
			   sprintf(temp,"投胶=%2d,投炭=%2d,投油=%2d,投小=%2d,转速=%2d,压力=%2d,段时=%2d,温度=%2d,能量=%5.2f,关系=%2d",
						p_y.huan[i].dou_jiao_time,p_y.huan[i].dou_tan_time,p_y.huan[i].dou_you_time,p_y.huan[i].dou_xiao_time,\
						p_y.huan[i].speed,p_y.huan[i].pre,p_y.huan[i].lian_time,p_y.huan[i].tem,p_y.huan[i].neg,p_y.huan[i].ctr);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }

			   sprintf(temp,"下温=%2d,上温=%2d,小重=%5.2f,排条=%2d,排时=%2d",
						p_y.base.min_wen_du,p_y.base.max_wen_du,\
						p_y.base.weight,p_y.base.pai_liao_ask,p_y.base.pai_liao_time);
					x=strlen(temp);j=j+20;
					TextOut(hdc,20,j,temp,x);
	//		   sprintf(temp,"节点名=%5d,数据=%2d,节点名=%5d,数据=%2d",
	//					elec_input.note_name[2],elec_input.data[2],elec_output.note_name[2],elec_output.data[2]);
	//				x=strlen(temp);j=j+20;
	//				TextOut(hdc,20,j,temp,x);
			   for(i=0;i<p_y.yt1_sum;i++)
			   {
			   sprintf(temp,"斗号=%s,次数=%2d,油2投1=%5.2f,快慢=%5.2f,点动=%5.2f,公差=%5.2f,控时=%3.2f,稳时=%2d",
						p_y.yt1[i].dou,p_y.yt1[i].add_time,p_y.yt1[i].weight,p_y.yt1[i].fast_do,\
						p_y.yt1[i].drop_do,p_y.yt1[i].gon_cha,p_y.yt1[i].control_time,p_y.yt1[i].stop_time);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }

		   ReleaseDC(hwnd,hdc);
		Sleep(1600);
		   write_multi_screen(200,60,temp);
        }
	   Sleep(330);
    }while(1);
        return FALSE;
}

// test dll
/*LRESULT WINAPI Write_Elec_Thread (int y)
{
    int k;
    int data=0;
    do{
         if (data==0) data=1;
         else   data=0;
          for(k=0;k<=elec_input.length;k++)
                        elec_input.data[k]=data;
          write_elec_input(&elec_input);

          for(k=0;k<=elec_output.length;k++)
                        elec_output.data[k]=data;
          write_elec_output(&elec_output);

          Sleep(990);
    }while(1);
	Sleep(2200);

    return FALSE;
}
*/

/*
 * FirstInstance - register window class for the application,
 *                 and do any other application initialization
 */
static BOOL FirstInstance( HANDLE this_inst )
{
    WNDCLASS    wc;
    BOOL        rc;

    /*
     * set up and register window class
     */
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
 * AnyInstance - do work required for every instance of the application:
 *                create the window, initialize data
 */
static BOOL AnyInstance( HANDLE this_inst, int cmdshow, LPSTR cmdline )
{

    extra_data  *edata_ptr;
    
    /*
     * create main window
     */
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

    /*
     * set up data associated with this window
     */
    edata_ptr = malloc( sizeof( extra_data ) );
    if( edata_ptr == NULL ) return( FALSE );
    edata_ptr->cmdline = cmdline;
    SetWindowLong( hwnd, EXTRA_DATA_OFFSET, (DWORD) edata_ptr );

    /*
     * display window
     */
    ShowWindow( hwnd, cmdshow );
    UpdateWindow( hwnd );
    
    return( TRUE );
                        
} /* AnyInstance */

/*
 * AboutDlgProc - processes messages for the about dialog.
 */

BOOL _EXPORT FAR PASCAL AboutDlgProc( HWND hwnd, unsigned msg,
                                UINT wparam, LONG lparam )
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
 * WindowProc - handle messages for the main application window
 */


/*LONG _EXPORT FAR PASCAL  ThreadWndProc(HWND hwnd,UINT msg,
                                     UINT wparam, LONG lparam )
{
	char temp[80];
	long i=0;
	HDC  hdc;
    switch( msg )
    {

      case WM_CREATE:
            SetForegroundWindow(hwnd);
          break;

    case WM_COMMAND:
        switch( LOWORD( wparam ) )
        {
            case IDM_LOCKUP:
                while (TRUE)
				{
					i++;
					hdc=GetDC(hwnd);
					sprintf(temp,"%8d",i);
			        TextOut(hdc,10,10,temp,8);
					ReleaseDC(hwnd,hdc);
					if(i>4000)
					{						
						HANDLE  hThread;
						
						hThread=GetCurrentThreadId();
						for(i=0;i<3;i++)
							if (chdthd[i]==hThread)
								break;
						SuspendThread(hThread);
						hdc=GetDC(hwnd);
						sprintf(temp,"%8s","睡眠 1s");
						TextOut(hdc,10,10,temp,8);
						ReleaseDC(hwnd,hdc);
						Sleep(1000);
						switch(i)
						{
						case 0:
							ResumeThread(chdthd[i+1]);
							ResumeThread(chdthd[i+2]);
							break;
						case 1:
							ResumeThread(chdthd[i+1]);
							ResumeThread(chdthd[i-1]);
							break;
						case 2:
							ResumeThread(chdthd[i-1]);
							ResumeThread(chdthd[i-2]);
							break;
						}	
						i=1;
					}
				}
				;
            break;

        case  IDM_EXIT:
                SendMessage(hwnd,WM_CLOSE,0,0);
        }
        break;

    case WM_DESTROY:
        PostQuitMessage( 0 );
        break;

    default:
        return( DefWindowProc( hwnd, msg, wparam, lparam ) );
    }
    return( 0L );

}
*/

LONG _EXPORT FAR PASCAL WindowProc( HWND hwnd, unsigned msg,
                                     UINT wparam, LONG lparam )
{
    FARPROC     proc;
    extra_data  *edata_ptr;
    char        temp[40];
    int i;

    DWORD    dwThreadID;


    switch( msg ) {
    case WM_CREATE:
	//	init_elec_input(&elec_input);
	//	init_elec_output(&elec_output);
/*	      i=sio_close(2);
          Sleep(220);
          i=sio_open(2);
		  if (i<0)
		  {
			   MessageBox(hwnd,"open error","open 2",0);               return  0;
		  }
            sio_ioctl(2,B9600, BIT_8 | STOP_1 | P_ODD);
//            wsprintf(temp, "call sio_ioctl return value = %d",sio_ioctl(2,
//				B9600, BIT_7 | STOP_1 | P_ODD));
			    sio_flush(2,1);	
	//creating second thread
			if (total_thread>0)
			{
				MessageBox(0,"已经产生了线程","",0);
				MessageBeep(0);	
				break;
			}		
			
		  hThread = CreateThread(NULL, 0,                            
                            (LPTHREAD_START_ROUTINE) Read_Elec_Thread,
                            (LPVOID)i,
                             0, &dwThreadID);       
		if  (hThread==0)
		{
			MessageBox(0,"create scond thread faile","",0);
				MessageBeep(0);		    
		}
		 chdthd[total_thread++]=hThread		;
          if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) Write_Elec_Thread,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
        chdthd[total_thread++]=hThread;
	*/
/*         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) Read_Run_Thread,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;
*/
         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) OneThread ,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;

/*
         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) TwoThread ,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;

         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) ThreeThread ,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;

         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) FourThread ,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;

  /*       if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) FiveThread ,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;
    */    break;
    case WM_COMMAND:
        switch( LOWORD( wparam ) ) 
		{
		case MENU_KILL:
		for(i=0;i<total_thread;i++)		 
		{			
			TerminateThread(chdthd[i],1);
		}
		total_thread=0;
		  i=sio_close(TEST_PORT);

		 break;
         
        case MENU_ABOUT:
            proc = MakeProcInstance( AboutDlgProc, MyInstance );
            DialogBox( MyInstance,"AboutBox", hwnd, proc );
            FreeProcInstance( proc );
            break;

        case MENU_CREATE:


        break;

/*	case MENU_COM_READ:
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
		 break;*/
		}//end case
		break;
    case WM_DESTROY:
		for(i=0;i<total_thread;i++)		 
		{			
			TerminateThread(chdthd[i],1);
		}
		total_thread=0;
		  i=sio_close(TEST_PORT);
		  if (i<0)
		  {
//			   MessageBox(hwnd,"close  error","open",0);
				return  0;
		  }
//		  Sio_Open_All( );

        PostQuitMessage( 0 );
        break;

    default:
        return( DefWindowProc( hwnd, msg, wparam, lparam ) );
    }
    return( 0L );

} /* WindowProc */

int h16_bcd(int i)
{
 int j,k,l;
 j=(int)(i/10.0);k=i-j*10;
 l=j;j=(int)(j/10.0);k=k+((l-10*j)<<4);
 l=j;j=(int)(j/10.0);k=k+((l-10*j)<<8);
 l=j;j=(int)(j/10.0);k=k+((l-10*j)<<12);
 return k;
}

int bcd_h16(int i)
{
 int j;
 j=1000*((i&0xf000)>>12);
 j=j+100*((i&0xf00)>>8);
 j=j+10*((i&0xf0)>>4);
 j=j+(i&0xf);
 printf(" j=%d \n",j);
 return j;
}

char *h16_b02(int i)
{
 int j=0;
 char p1[6];
 j=((i&0x8)<<9);
 j=j+((i&0x4)<<6);
 j=j+((i&0x2)<<3);
 j=j+(i&0x1);
 sprintf(p1,"%x",j);
// strcpy(d_3,p1);
 return p1;
}

char *ii_bin(unsigned int i,char *p)
{
 char p1[7];
 long j;
 j=i+10000;*p=0;
 sprintf(p1,"%ld",j);
 strcpy(p,p1+1);
 return p;
}

char *i_bin(unsigned int i,char *p)
{
 char p1[7];
 sprintf(p1,"%d",i);
 strcpy(p,p1);
 return p;
}

char *h16_16a(unsigned int i,char *p)
{
 int k,l;
 long j;
 char p1[10];
 j=i+0x10000;
 *p=0;
 sprintf(p1,"%lx",j);
 //strupr[p1];
 strcpy(p,p1+1);
 l=strlen(p);
 for(k=0;k<l;k++)
  if(*(p+k)>0x60) *(p+k)=*(p+k)-0x20;
 return p;
}

char *head_d(char *p,char *p1,unsigned int m,unsigned int i)
{
 int j,k;
 char p2[7];
 strcat(p,p1);
 j=m+1000000;
 sprintf(p2,"%d",j);
 strcat(p,p2+1);
 *p2=0;
 j=i+0x100;
 sprintf(p2,"%x",j);
 for(k=0;k<3;k++)
  if(*(p2+k)>0x60) *(p2+k)=*(p2+k)-0x20;
 strcat(p,p2+1);
 return p;
}

char *head_h(char *p,char *p1,unsigned int m,unsigned int i)
{
 int j,k;
 char p2[7];
 strcat(p,p1);
 j=m+0x1000000;
 sprintf(p2,"%x",j);
 strcat(p,p2+1);
 *p2=0;
 j=i+0x100;
 sprintf(p2,"%x",j);
 for(k=0;k<3;k++)
  if(*(p2+k)>0x60) *(p2+k)=*(p2+k)-0x20;
 strcat(p,p2+1);
 p=strupr(p);
 return p;
}

/*void head_d(char *p,char *p1,unsigned int i)
{
 int j,k,l;
 char p2[10],p3[10];
 l=strlen(p1);
 for(j=0;j<5;j++) p2[j]='0';
 j=0;
 while(p1[j]>'9') {
  p2[j]=p1[j];
  j++;
 }
 for(k=l;k>j;k--)
  p2[5-k+j]=p1[j+l-k];
 p2[5]=0;
 strcat(p,p2);
 *p2=0;*p3=0;
 j=i+100;
 sprintf(p2,"%d",j);
 strcpy(p3,p2+1);
 strcat(p,p3);
// return p;
}*/

unsigned asc_16i(char *p)
{
 int i,l[5];
 for(i=0;i<4;i++)
  if(p[i]>'9') l[i]=(p[i]-0x37)&0xdf; else \
              l[i]=(p[i]-0x30)&0xdf;
 return(16*16*16*l[0]+16*16*l[1]+16*l[2]+l[3]);
}

unsigned asc_10i(char *p)
{
 return(10*10*10*(p[0]-0x30)+10*10*(p[1]-0x30)+\
        10*(p[2]-0x30)+(p[3]-0x30));
}

char *ad_d(char *p)
{
 int i,j;
 char p1;
 p1=0;
 j=strlen(p);
/* for(i=1;i<j;i++) p1=p1+(*(p+i));
 if((p1&0xf0)>0x90) *(p+i)=0x37+((p1&0xf0)>>4);
 else *(p+i)=0x30+((p1&0xf0)>>4);
 if((p1&0xf)>0x9) *(p+i+1)=0x37+(p1&0xf);
 else *(p+i+1)=0x30+(p1&0xf);
 *(p+i+2)=0;*/
 *(p+j)=0;
 return(p);
}

char *head_t(char *p,char *p1)
{
 *p=0x5;*(p+1)='0';*(p+2)='0';*(p+3)='F';*(p+4)='F';*(p+5)=0;
 strcat(p,p1);
 *(p+7)='0';*(p+8)=0;
 return p;
}
void bcom1(char *p)
{
	int j,k;
//	HDC  hdc;
//	hdc=GetDC(hwnd);
//	TextOut(hdc,10,50,d_1,strlen(d_1));
//	ReleaseDC(hwnd,hdc);
//	sio_putch(2,0x04);
//	Sleep(20);
	sio_flush(2,1);	
	sio_putch(2,0x0c);
	Sleep(2);
	j=strlen(p);
	for(k=0;k<j;k++)
	{
	 sio_putch(2,*(p+k));
//	 sio_write(2,*(p+k),1);
	 Sleep(2);
	}
//	sio_write(2,d_1,strlen(d_1));
}

void bcom2()
{
 int i,j,k;
 char temp[20];
// HDC  hdc;
// hdc=GetDC(hwnd);
// sio_flush(2,0);	
// Sleep(50);
// sio_putch(2,0x0c);
// Sleep(60);
// for(i=0;i<8;i++)
//     temp[i]=' ';
//    temp[i]='\0';

 i=0;
 do{
//   sio_read(2,temp,1);
//	temp[i]=temp[0]+0x30;
//	temp[i]=j+0x30;
//    TextOut(hdc,10,100,temp,9);
	j=sio_getch(2);
//	Sleep(1);
   }while(j != 2);
 i=0;
 do{
    j=sio_getch(2);
	if((j&0xff00)==0)  d_2[i]=j;
//    sio_read(2,temp,1);
//    if((temp[0]&0xff00)==0)  d_2[i]=temp[0];
    i++;
	Sleep(1);
   }while(d_2[i-1] != 0x03);
  d_2[i-1]=0;
	sio_putch(2,0x06);
// TextOut(hdc,10,150,d_2,strlen(d_2));
 j=strlen(d_2);
 k=(int)(j/4);
 for(i=1;i<k;i++)
  for(j=0;j<4;j++)
  {p_d[i][j]=d_2[i*4+j];
   p_d[i][j+1]=0;
//   TextOut(hdc,200,(80+i*20),p_d[i],4);
  }
 for(k=1;k<i;k++)
  d_d[k]=asc_16i(p_d[k]);
	sio_flush(2,1);	
// Sleep(2);
// ReleaseDC(hwnd,hdc);
 //return k;
}
void bcom3()
{
 int i,j,k;
 char temp[20];
// HDC  hdc;

// hdc=GetDC(hwnd);
 for(i=0;i<8;i++)
     temp[i]=' ';
    temp[i]='\0';

 i=0;
 do{
	j=sio_getch(2);

   }while(j != 2);
 i=0;
 do{
    j=sio_getch(2);
	if((j&0xff00)==0)  d_2[i]=j;
    i++;
	Sleep(1);
   }while(d_2[i-1] != 0x03);
 d_2[i-1]=0;
// TextOut(hdc,10,250,d_2,strlen(d_2));
 j=strlen(d_2);
 k=(int)(j/4);
 for(i=1;i<k;i++)
  for(j=0;j<4;j++)
  {p_d[i][j]=d_2[i*4+j];
   p_d[i][j+1]=0;
  }
 for(k=1;k<i;k++)
	   d_d[k]=asc_10i(p_d[k]);
	sio_flush(2,1);	
// Sleep(2);
// ReleaseDC(hwnd,hdc);
 //return k;
}

void bcom4()
{
 int i,j;
 char temp[20];
 HDC  hdc;

 hdc=GetDC(hwnd);
 for(i=0;i<8;i++)
     temp[i]=' ';
    temp[i]='\0';

 i=0;
 do{
	j=sio_getch(2);

   }while(j != 2);
 i=0;
 do{
    j=sio_getch(2);
	if((j&0xff00)==0)  d_2[i]=j;
    i++;
//	Sleep(2);
   }while(d_2[i-1] != 0x03);
 d_2[i-1]=0;
// TextOut(hdc,10,220,d_2,strlen(d_2));
 d_3[0]=d_2[19];
 d_3[1]=d_2[17];
 d_3[2]=d_2[23];
 d_3[3]=d_2[21];
 d_3[4]='/0';
 d_d[0]=asc_10i(d_3);
 d_f=d_d[0]/100.0;
 sio_flush(2,1);	
 //Sleep(2);
 ReleaseDC(hwnd,hdc);
 //return k;
}
void read_D()
{
	int j;
	head_t(d_1,"QR");
	head_d(d_1,"D",600,10);
//	ad_d(d_1);
//	TextOut(hdc,10,300,d_1,strlen(d_1));
	bcom1(d_1);
	bcom2();//(二十进制)用bcom3()
	for(j=0;j<10;j++)
	{
		data_d[j]=d_d[j+1];
	}
	Sleep(110);
}
void read_X()
{
	int j;
	j=0xc0;
	head_t(d_1,"QR");
	head_h(d_1,"X",j,8);
	ad_d(d_1);
//		TextOut(hdc,10,300,d_1,strlen(d_1));
	bcom1(d_1);
	bcom2();
	for(j=0;j<16;j++)
		if((d_d[1]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j,1);
//			elec_input.data[i*32+j]=1;
		else
			Set_Light(2192+j,0);
//			elec_input.data[i*32+j]=0;
	for(j=0;j<16;j++)
		if((d_d[2]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j+16,1);
		else
			Set_Light(2192+j+16,0);
	for(j=0;j<16;j++)
		if((d_d[3]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j+32,1);
		else
			Set_Light(2192+j+32,0);
	for(j=0;j<16;j++)
		if((d_d[4]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j+48,1);
		else
			Set_Light(2192+j+48,0);
	for(j=0;j<16;j++)
		if((d_d[5]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j+64,1);
//			elec_input.data[i*32+j]=1;
		else
			Set_Light(2192+j+64,0);
//			elec_input.data[i*32+j]=0;
	for(j=0;j<16;j++)
		if((d_d[6]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j+80,1);
		else
			Set_Light(2192+j+80,0);
	for(j=0;j<16;j++)
		if((d_d[7]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j+96,1);
		else
			Set_Light(2192+j+96,0);
	for(j=0;j<16;j++)
		if((d_d[8]&(unsigned short)pow(2,j))>0)
			Set_Light(2192+j+112,1);
		else
			Set_Light(2192+j+112,0);
}
void read_Y()
{
	int j;
	j=0x140;
	head_t(d_1,"QR");
	head_h(d_1,"Y",j,11);
	ad_d(d_1);
	bcom1(d_1);
	bcom2();
	for(j=0;j<16;j++)
		if((d_d[1]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j,1);
		else
			Set_Turn(2320+j,0);
	for(j=0;j<16;j++)
		if((d_d[2]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+16,1);
		else
			Set_Turn(2320+j+16,0);
	for(j=0;j<16;j++)
		if((d_d[3]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+32,1);
		else
			Set_Turn(2320+j+32,0);
	for(j=0;j<16;j++)
		if((d_d[4]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+48,1);
		else
			Set_Turn(2320+j+48,0);
	for(j=0;j<16;j++)
		if((d_d[5]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+64,1);
		else
			Set_Turn(2320+j+64,0);
	for(j=0;j<16;j++)
		if((d_d[6]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+80,1);
		else
			Set_Turn(2320+j+80,0);
	for(j=0;j<16;j++)
		if((d_d[7]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+96,1);
		else
			Set_Turn(2320+j+96,0);
	for(j=0;j<16;j++)
		if((d_d[8]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+112,1);
		else
			Set_Turn(2320+j+112,0);
	for(j=0;j<16;j++)
		if((d_d[9]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+128,1);
		else
			Set_Turn(2320+j+128,0);
	for(j=0;j<16;j++)
		if((d_d[10]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+144,1);
		else
			Set_Turn(2320+j+144,0);
	for(j=0;j<16;j++)
		if((d_d[11]&(unsigned short)pow(2,j))>0)
			Set_Turn(2320+j+160,1);
		else
			Set_Turn(2320+j+160,0);
}
void read_M1()
{
	int j;
	j=288;
	head_t(d_1,"QR");
	head_d(d_1,"M",j,4);//从M288点开始读到M352点
	ad_d(d_1);
	bcom1(d_1);
	bcom2();
//	sprintf(temp,"d_d=%3x",d_d[1]);
//	TextOut(hdc,10,300,temp,10);
	for(j=0;j<16;j++)
		if((d_d[1]&(unsigned short)pow(2.0,j))==0.0)
			data_m1[j]=0;
				else
			data_m1[j]=1;
	for(j=0;j<16;j++)
		if((d_d[2]&(unsigned short)pow(2.0,j))==0.0)
			data_m1[j+16]=0;
		else
			data_m1[j+16]=1;
	for(j=0;j<16;j++)
		if((d_d[3]&(unsigned short)pow(2.0,j))==0.0)
			data_m1[32+j]=0;
		else
			data_m1[32+j]=1;
	for(j=0;j<16;j++)
		if((d_d[4]&(unsigned short)pow(2.0,j))==0.0)
			data_m1[j+48]=0;
		else
			data_m1[j+48]=1;
}
void read_M2()
{
	int j;
	j=288;
	head_t(d_1,"QR");
	head_d(d_1,"M",j,3);//从M592点开始读到M640点
	ad_d(d_1);
	bcom1(d_1);
	bcom2();
//	sprintf(temp,"d_d=%3x",d_d[1]);
//	TextOut(hdc,10,300,temp,10);
	for(j=0;j<16;j++)
		if((d_d[1]&(unsigned short)pow(2.0,j))==0.0)
			data_m1[j+64]=0;
		else
			data_m1[j+64]=1;
	for(j=0;j<16;j++)
		if((d_d[2]&(unsigned short)pow(2.0,j))==0.0)
			data_m1[j+80]=0;
		else
			data_m1[j+80]=1;
	for(j=0;j<16;j++)
		if((d_d[3]&(unsigned short)pow(2.0,j))==0.0)
			data_m1[j+96]=0;
		else
			data_m1[j+96]=1;
}
void read_D_end()
{
	int i,j;
	for(i=0;i<3;i++)
		{
		head_t(d_1,"QR");
		head_d(d_1,"D",600+i*20,20);
//		ad_d(d_1);
//		TextOut(hdc,10,300,d_1,strlen(d_1));
		bcom1(d_1);
		bcom2();//PLC中10进制改为16进制时改用bcom3()读出
		for(j=0;j<20;j++)
			{
			data_d[i*20+j]=d_d[j+1];
//			f3=(float)(data_d[59]/100.0); //i*2.9;
//				sprintf(temp,"data_d=%4d f3=%f",data_d[59],f3);
//				TextOut(hdc,400,20,temp,20);
			}
		}
}
void write_Y()
{
	int i,j;
	unsigned int y_1;
	char zoj[20];
	short data_y[181];
 HDC  hdc;
 hdc=GetDC(hwnd);
	init_elec_output(&elec_output);
	for(i=0;i<elec_output.length;i++)
	{
		if(elec_output.note_name[i]>2319 && elec_output.note_name[i]<2500)
			data_y[elec_output.note_name[i]-2320]=elec_output.data[i];
	}
	for(i=0;i<6;i++)
	{
		j=0x140+i*0x20;
		head_t(d_1,"QW");
		head_h(d_1,"Y",j,2);
		y_1=0;
		for(j=0;j<16;j++)
			if(data_y[i*32+j]==1) y_1=y_1+(unsigned short)pow(2,j);
			strcat(d_1,h16_16a(y_1,zoj));
			y_1=0;
		for(j=0;j<16;j++)
			if(data_y[i*32+j+16]==1) y_1=y_1+(unsigned short)pow(2,j);
			strcat(d_1,h16_16a(y_1,zoj));
			ad_d(d_1);
//			TextOut(hdc,10,80,d_1,strlen(d_1)+1);
    		bcom1(d_1);
			Sleep(110);
	 }
 ReleaseDC(hwnd,hdc);
}
void write_p_f()
{
	int i,j;
	char zoj[120];
	char th1_dh[66],th2_dh[66],yz1_dh[66],yz2_dh[66];
	read_produce(&p_y );
//	init_elec_input(&elec_input);
//	init_elec_output(&elec_output);
	th1_dh[0]=0;th2_dh[0]=0;yz1_dh[0]=0;yz2_dh[0]=0;
	for(i=p_y.th1_sum;i>=0;i--)
		strcat(th1_dh,p_y.th1[i].dou);
	for(i=p_y.th2_sum;i>=0;i--)
		strcat(th2_dh,p_y.th2[i].dou);
 	head_t(d_1,"QW");
	head_d(d_1,"D",300,5);
	if((p_y.th1_sum!=0)||(p_y.th2_sum!=0))
	{
	if(p_y.th2_sum==0)
	strcat(d_1,h16_16a(1,zoj));
	else
	strcat(d_1,h16_16a(2,zoj));
	}
	else
	//PLC中16进制改为10进制把ii_bin改成h16_16a
	strcat(d_1,h16_16a(0,zoj));
	strcat(d_1,h16_16a(p_y.th1_sum,zoj));
	strcat(d_1,h16_16a(atoi(th1_dh),zoj));
	strcat(d_1,h16_16a(p_y.th2_sum,zoj));
	strcat(d_1,h16_16a(atoi(th2_dh),zoj));
	ad_d(d_1);
//	TextOut(hdc,10,100,d_1,strlen(d_1));
	bcom1(d_1);
	for(i=p_y.yz1_sum;i>=0;i--)
		strcat(yz1_dh,p_y.yz1[i].dou);
	for(i=p_y.yz2_sum;i>=0;i--)
		strcat(yz2_dh,p_y.yz2[i].dou);
 	head_t(d_1,"QW");
	head_d(d_1,"D",305,5);
	if((p_y.yz1_sum!=0)||(p_y.yz2_sum!=0))
	{
	if(p_y.yz2_sum==0)
	strcat(d_1,h16_16a(1,zoj));
	else
	strcat(d_1,h16_16a(2,zoj));
	}
	else
	strcat(d_1,h16_16a(0,zoj));
	strcat(d_1,h16_16a(p_y.yz1_sum,zoj));
	strcat(d_1,h16_16a(atoi(yz1_dh),zoj));
	strcat(d_1,h16_16a(p_y.yz2_sum,zoj));
	strcat(d_1,h16_16a(atoi(yz2_dh),zoj));
	ad_d(d_1);
//	TextOut(hdc,10,200,d_1,strlen(d_1));
	bcom1(d_1);
	for(i=0;i<7;i++)
	{
//   sprintf(temp,"重量=%5.2f,公差=%5.2f,快慢=%5.2f,点动=%5.2f,斗号=%s",
//			p_y.th1[0].weight,p_y.th1[0].gon_cha,p_y.th1[0].fast_do,p_y.th1[0].drop_do,p_y.th1[0].dou);
//			x=strlen(temp);
//			TextOut(hdc,20,150,temp,x);
		if(p_y.th1[i].weight>0.0)
		{
			head_t(d_1,"QW");
			j=300+(atoi(p_y.th1[i].dou)*10);
			head_d(d_1,"D",j,5);
			strcat(d_1,h16_16a((int)(p_y.th1[i].weight*10),zoj));//PLC中16进制改为10进制把h16_16a改成h16_16a
		//	strcat(d_1,ii_bin(0,zoj));
			strcat(d_1,h16_16a((int)(p_y.th1[i].fast_do*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.th1[i].drop_do*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.th1[i].control_time*10.01),zoj));
		//	strcat(d_1,ii_bin((int)(p_y.th1[i].stop_time*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.th1[i].gon_cha*10.01),zoj));
			ad_d(d_1);
			bcom1(d_1);
		}
		if(i==0)
		{
			head_t(d_1,"QW");
			head_d(d_1,"D",550,1);
			strcat(d_1,h16_16a((int)(p_y.th1[i].stop_time*10),zoj));
			ad_d(d_1);
			bcom1(d_1);
		}
	}
	for(i=0;i<7;i++)
	{
		if(p_y.th2[i].weight>0.0)
		{
			head_t(d_1,"QW");
			j=305+(atoi(p_y.th2[i].dou)*10);
			head_d(d_1,"D",j,5);
			strcat(d_1,h16_16a((int)(p_y.th2[i].weight*10),zoj));
		//	strcat(d_1,ii_bin(0,zoj));
			strcat(d_1,h16_16a((int)(p_y.th2[i].fast_do*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.th2[i].drop_do*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.th2[i].control_time*10.01),zoj));
		//	strcat(d_1,ii_bin((int)(p_y.th2[i].stop_time*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.th2[i].gon_cha*10.01),zoj));
			ad_d(d_1);
			bcom1(d_1);
		}
	}

	for(i=0;i<3;i++)
	{
		if(p_y.yz1[i].weight>0.0)
		{
			head_t(d_1,"QW");
			j=370+(atoi(p_y.yz1[i].dou)*10);
			head_d(d_1,"D",j,5);
			strcat(d_1,h16_16a((int)(p_y.yz1[i].weight*100.01),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz1[i].fast_do*100.01),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz1[i].drop_do*100.01),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz1[i].control_time*10.01),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz1[i].gon_cha*100.01),zoj));
			ad_d(d_1);
			bcom1(d_1);
		}
		if(i==0)
		{
			head_t(d_1,"QW");
			head_d(d_1,"D",551,1);
			strcat(d_1,h16_16a((int)(p_y.yz1[i].stop_time*10),zoj));
			ad_d(d_1);
			bcom1(d_1);
		}
	}
	for(i=0;i<3;i++)
	{
		if(p_y.yz2[i].weight>0.0)
		{
			head_t(d_1,"QW");
			j=375+(atoi(p_y.yz2[i].dou)*10);
			head_d(d_1,"D",j,5);
			strcat(d_1,h16_16a((int)(p_y.yz2[i].weight*100),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz2[i].fast_do*100),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz2[i].drop_do*100),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz2[i].control_time*10.01),zoj));
	//		strcat(d_1,ii_bin((int)(p_y.yz2[i].stop_time*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.yz2[i].gon_cha*100.01),zoj));
			ad_d(d_1);
			bcom1(d_1);
		}
	}
	head_t(d_1,"QW");
	head_d(d_1,"D",410,3);
	if((p_y.sj1_sum!=0)||(p_y.sj2_sum!=0))
	{
	if(p_y.sj2_sum==0)
	strcat(d_1,h16_16a(1,zoj));
	else
	strcat(d_1,h16_16a(2,zoj));
	}
	else
	strcat(d_1,h16_16a(0,zoj));
	strcat(d_1,h16_16a(p_y.sj1_sum,zoj));
//	if(p_y.sj2_sum==0)
//	strcat(d_1,ii_bin(0,zoj));
//	else
	strcat(d_1,h16_16a(p_y.sj2_sum,zoj));
	ad_d(d_1);
	bcom1(d_1);
	for(i=0;i<p_y.sj1_sum;i++)
	{
		if(p_y.sj1[i].weight>0.0)
		{
			head_t(d_1,"QW");
			j=415+(i*3);
			head_d(d_1,"D",j,3);
			strcat(d_1,h16_16a((int)(p_y.sj1[i].weight*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.sj1[i].gon_cha*10.01),zoj));
			strcat(d_1,h16_16a(p_y.sj1[i].stop_time,zoj));	//胶状
			ad_d(d_1);
			bcom1(d_1);
		}
	}
	for(i=0;i<p_y.sj2_sum;i++)
	{
		if(p_y.sj2[i].weight>0.0)
		{
			head_t(d_1,"QW");
			j=430+(i*3);
			head_d(d_1,"D",j,3);
			strcat(d_1,h16_16a((int)(p_y.sj2[i].weight*10),zoj));
			strcat(d_1,h16_16a((int)(p_y.sj2[i].gon_cha*10.01),zoj));
			strcat(d_1,h16_16a(p_y.sj2[i].stop_time,zoj));	//胶状
			ad_d(d_1);
			bcom1(d_1);
		}
	}
	head_t(d_1,"QW");
	head_d(d_1,"D",440,5);
	strcat(d_1,h16_16a(p_y.base.max_wen_du,zoj));
	strcat(d_1,h16_16a(p_y.base.min_wen_du,zoj));
	strcat(d_1,h16_16a(p_y.base.pai_liao_time,zoj));
	strcat(d_1,h16_16a(p_y.base.pai_liao_ask,zoj));
	strcat(d_1,h16_16a((int)(p_y.base.weight*100),zoj));
	ad_d(d_1);
	bcom1(d_1);
	head_t(d_1,"QW");
	head_d(d_1,"D",445,5);
	strcat(d_1,h16_16a(0,zoj));
	strcat(d_1,h16_16a(p_y.base.temp1,zoj));
	strcat(d_1,h16_16a(p_y.pei_max,zoj));
	strcat(d_1,h16_16a(run_table_do.total_che,zoj));
	strcat(d_1,h16_16a(p_y.huan[0].speed,zoj));
	ad_d(d_1);
	bcom1(d_1);
	for(i=0;i<p_y.pei_max;i++)
	{
	j=450+(i*10);
	head_t(d_1,"QW");
	head_d(d_1,"D",j,5);
	strcat(d_1,h16_16a(p_y.huan[i].dou_jiao_time,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].dou_tan_time,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].dou_you_time,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].dou_xiao_time,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].st,zoj));
	ad_d(d_1);
	bcom1(d_1);
	head_t(d_1,"QW");
	head_d(d_1,"D",j+5,5);
	strcat(d_1,h16_16a(p_y.huan[i].pre,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].lian_time,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].tem,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].neg,zoj));
	strcat(d_1,h16_16a(p_y.huan[i].ctr,zoj));
	ad_d(d_1);
	bcom1(d_1);
	}
}
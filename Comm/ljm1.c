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

char d_1[40],d_2[40],p_d[10][6];
char d_3[6],d_4[6];
int  d_d[10],cs_flag=0;
int	 data_d[70],data_m1[65],data_m2[65];
float d_f;
	struct  elect_send_type     elec_input;  //elec
	struct  elect_send_type     elec_output;  //elec
	struct  make_lian_liao Duan_Data;
	struct  make_base_type Make_Temp;


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
	int i=0;

    MyInstance = this_inst;
#ifdef __WINDOWS_386__
    sprintf( GenericClass,"GenericClass%d", this_inst );
    prev_inst = 0;
#endif
	hsem=CreateSemaphore(NULL,1,1,"ZGH");
	if (GetLastError()==ERROR_ALREADY_EXISTS)
	{
		MessageBox(0,"已经运行 ","错误",0);		
		CloseHandle(hsem);
		return FALSE;
	}
	
    if( !prev_inst ) 
	{
        if( !FirstInstance( this_inst ) ) return( FALSE );
    }
	
    if( !AnyInstance( this_inst, cmdshow, cmdline ) ) return( FALSE );

    while( GetMessage( &msg, NULL, NULL, NULL ) ) {

        TranslateMessage( &msg );
        DispatchMessage( &msg );

    }

		sio_close(TEST_PORT);
		for(i=0;i<total_thread;i++)		 
		{			
			TerminateThread(chdthd[i],1);
		}
		total_thread=0;
        PostQuitMessage( 0 );
    
     return FALSE;

    return( msg.wParam );

} /* WinMain */


LRESULT WINAPI write_Thread (int x)
{
     char  temp[TEST_LEN+1];
     int i;
	 HDC   hdc;
		i=0;
		  x=sio_open(TEST_PORT);
		  if (x<0)
		  {
			   //MessageBox(hwnd,"open  error","open",0);
				return  0;
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


}



LRESULT WINAPI OneThread (int x)
{
	int i=0,j,k,t_1=0,t_2=0;
	char *s1,*s2;
	char zoj[20];
	char temp[40];

	float f1,f2,f3,f4;
	HDC  hdc;
    while (TRUE)
		{
        hdc=GetDC(hwnd);
		do
		{
			Sleep(220);
		}while(cs_flag==0);
//		run_table_do.all_che_time=1;
//		run_table_do.duan_time=1;
//		write_run(&run_table_do,-1);
		f1=9.8;
		for(i=0;i<10;i++)
		{
			f2=i*8.2;
			f3=i*2.9;
			f4=i*12.6;
		write_ban(1,1,&f1,&f2,2);
		write_ban(1,2,&f1,&f3,2);
		write_ban(1,4,&f1,&f4,2);
/*                elec_output.data[i]=1;
                  write_elec_output(&elec_output);
*/
//		write_power(i,i*1.3,i*5,1);
//		run_table_do.now_che=i;
//		write_run(&run_table_do,-1);
                write_error(355,"DFASDF");
                write_error(1,"测试错误型号");
		Sleep(990);
		}
//		run_table_do.duan_time=0;
//		run_table_do.all_che_time=0;
//		write_run(&run_table_do,-1);
        ReleaseDC(hwnd,hdc);
        Sleep(1100);
	}

}

LRESULT WINAPI TwoThread (int x)
{
 /*   char  temp[40],temp1[40];
    int i=0,j,k;
	float f1;
	char s1[40],s2[40];
	unsigned c1;

	HDC  hdc;
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
            return FALSE;

}
	long tim1,tim2;
	time_t tim_1,tim_2,tim_3;

LRESULT WINAPI ThreeThread (int x)
{
	int i,j,k,t_1=0,t_2,jt=0;
	char *s1,*s2;
	char zoj[20];
	char temp[80];
	char th1_dh[8],th2_dh[8],yz1_dh[8],yz2_dh[8];
	char D_sx[4][9];
//	struct  elect_send_type     elec_input;  //elec
//	struct  elect_send_type     elec_output;  //elec
		  HDC  hdc;
	      init_elec_input(&elec_input);
		  init_elec_output(&elec_output);
		  cs_flag=0;
	 /*     i=sio_close(2);
          Sleep(220);
          i=sio_open(2);
		  if (i<0)
		  {
//			   MessageBox(hwnd,"open error","open 2",0);
               return  0;
		  }
            wsprintf(temp, "call sio_ioctl return value = %d", sio_ioctl(2,
				B9600, BIT_7 | STOP_1 | P_ODD));
			    sio_flush(2,1);	
                Sleep(100);
*/
    
 while (TRUE)
	{
     hdc=GetDC(hwnd);
	 do
	 {
      read_run(&run_table_do, &run_flag);
	  Sleep(1100);
     }while(run_flag==STOP_DOING);
/*******************************************************************
form  here 
	update  by zgh  98.4.2  
		error on run_flag=3
*****************************************************************
	I change  run_flag. 	but I don't konw cs_flag meaning 
*******************************************************************/
	 sprintf(temp,"运行开始  ");
	 x=strlen(temp);
	 TextOut(hdc,30,50,temp,x);
	 while(run_table_do.now_che<run_table_do.total_che)
	 {
		if ((run_flag==1 || run_flag==2) && cs_flag==0)
        {
		read_produce(&p_y);
		init_elec_input(&elec_input);
		init_elec_output(&elec_output);
		th1_dh[0]=0;th2_dh[0]=0;yz1_dh[0]=0;yz2_dh[0]=0;
		for(i=p_y.th1_sum;i>=0;i--)
			strcat(th1_dh,p_y.th1[i].dou);
		for(i=p_y.th2_sum;i>=0;i--)
			strcat(th2_dh,p_y.th2[i].dou);
	 	head_t(d_1,"WW");
		head_d(d_1,"D",300,5);
		if((p_y.th1_sum!=0)||(p_y.th2_sum!=0))
		{
		if(p_y.th2_sum==0)
		strcat(d_1,ii_bin(1,zoj));
		else
		strcat(d_1,ii_bin(2,zoj));
		}
		else
		strcat(d_1,ii_bin(0,zoj));
		strcat(d_1,ii_bin(p_y.th1_sum,zoj));
		strcat(d_1,ii_bin(atoi(th1_dh),zoj));
		strcat(d_1,ii_bin(p_y.th2_sum,zoj));
		strcat(d_1,ii_bin(atoi(th2_dh),zoj));

		ad_d(d_1);
    
		bcom1(d_1);
		for(i=p_y.yz1_sum;i>=0;i--)
			strcat(yz1_dh,p_y.yz1[i].dou);
		for(i=p_y.yz2_sum;i>=0;i--)
			strcat(yz2_dh,p_y.yz2[i].dou);
	 	head_t(d_1,"WW");
		head_d(d_1,"D",305,5);
		if((p_y.yz1_sum!=0)||(p_y.yz2_sum!=0))
		{
		if(p_y.yz2_sum==0)
		strcat(d_1,ii_bin(1,zoj));
		else
		strcat(d_1,ii_bin(2,zoj));
		}
		else
		strcat(d_1,ii_bin(0,zoj));
		strcat(d_1,ii_bin(p_y.yz1_sum,zoj));
		strcat(d_1,ii_bin(atoi(yz1_dh),zoj));
		strcat(d_1,ii_bin(p_y.yz2_sum,zoj));
		strcat(d_1,ii_bin(atoi(yz2_dh),zoj));
		ad_d(d_1);

		bcom1(d_1);
		for(i=0;i<7;i++)
		{
			if(p_y.th1[i].weight>0.0)
			{
				head_t(d_1,"WW");
				j=300+(atoi(p_y.th1[i].dou)*10);
				head_d(d_1,"D",j,5);
				strcat(d_1,ii_bin((int)(p_y.th1[i].weight*10),zoj));
			//	strcat(d_1,ii_bin(0,zoj));
				strcat(d_1,ii_bin((int)(p_y.th1[i].fast_do*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.th1[i].drop_do*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.th1[i].control_time*10.01),zoj));

				strcat(d_1,ii_bin((int)(p_y.th1[i].gon_cha*10.01),zoj));
				ad_d(d_1);
				bcom1(d_1);
			}
			if(i==0)
			{
				head_t(d_1,"WW");
				head_d(d_1,"D",550,1);
				strcat(d_1,ii_bin((int)(p_y.th1[i].stop_time*10),zoj));
				ad_d(d_1);
				bcom1(d_1);
			}
		}
		for(i=0;i<7;i++)
		{
			if(p_y.th2[i].weight>0.0)
			{
				head_t(d_1,"WW");
				j=305+(atoi(p_y.th2[i].dou)*10);
				head_d(d_1,"D",j,5);
				strcat(d_1,ii_bin((int)(p_y.th2[i].weight*10),zoj));
			//	strcat(d_1,ii_bin(0,zoj));
				strcat(d_1,ii_bin((int)(p_y.th2[i].fast_do*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.th2[i].drop_do*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.th2[i].control_time*10.01),zoj));
			//	strcat(d_1,ii_bin((int)(p_y.th2[i].stop_time*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.th2[i].gon_cha*10.01),zoj));
				ad_d(d_1);
				bcom1(d_1);
			}
		}

		for(i=0;i<3;i++)
		{
			if(p_y.yz1[i].weight>0.0)
			{
				head_t(d_1,"WW");
				j=370+(atoi(p_y.yz1[i].dou)*10);
				head_d(d_1,"D",j,5);
				strcat(d_1,ii_bin((int)(p_y.yz1[i].weight*100.01),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz1[i].fast_do*100.01),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz1[i].drop_do*100.01),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz1[i].control_time*10.01),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz1[i].gon_cha*100.01),zoj));
				ad_d(d_1);
				bcom1(d_1);
			}
			if(i==0)
			{
				head_t(d_1,"WW");
				head_d(d_1,"D",551,1);
				strcat(d_1,ii_bin((int)(p_y.yz1[i].stop_time*10),zoj));
				ad_d(d_1);
				bcom1(d_1);
			}
		}
		for(i=0;i<3;i++)
		{
			if(p_y.yz2[i].weight>0.0)
			{
				head_t(d_1,"WW");
				j=375+(atoi(p_y.yz2[i].dou)*10);
				head_d(d_1,"D",j,5);
				strcat(d_1,ii_bin((int)(p_y.yz2[i].weight*100),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz2[i].fast_do*100),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz2[i].drop_do*100),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz2[i].control_time*10.01),zoj));
		//		strcat(d_1,ii_bin((int)(p_y.yz2[i].stop_time*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.yz2[i].gon_cha*100.01),zoj));
				ad_d(d_1);
				bcom1(d_1);
			}
		}
		head_t(d_1,"WW");
		head_d(d_1,"D",410,3);
		if((p_y.sj1_sum!=0)||(p_y.sj2_sum!=0))
		{
		if(p_y.sj2_sum==0)
		strcat(d_1,ii_bin(1,zoj));
		else
		strcat(d_1,ii_bin(2,zoj));
		}
		else
		strcat(d_1,ii_bin(0,zoj));
		strcat(d_1,ii_bin(p_y.sj1_sum,zoj));
//		if(p_y.sj2_sum==0)
//		strcat(d_1,ii_bin(0,zoj));
//		else
		strcat(d_1,ii_bin(p_y.sj2_sum,zoj));
		ad_d(d_1);
		bcom1(d_1);
		for(i=0;i<p_y.sj1_sum;i++)
		{
			if(p_y.sj1[i].weight>0.0)
			{
				head_t(d_1,"WW");
				j=415+(i*3);
				head_d(d_1,"D",j,3);
				strcat(d_1,ii_bin((int)(p_y.sj1[i].weight*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.sj1[i].gon_cha*10.01),zoj));
				strcat(d_1,ii_bin(p_y.sj1[i].stop_time,zoj));	//胶状
				ad_d(d_1);
				bcom1(d_1);
			}
		}
		for(i=0;i<p_y.sj2_sum;i++)
		{
			if(p_y.sj2[i].weight>0.0)
			{
				head_t(d_1,"WW");
				j=430+(i*3);
				head_d(d_1,"D",j,3);
				strcat(d_1,ii_bin((int)(p_y.sj2[i].weight*10),zoj));
				strcat(d_1,ii_bin((int)(p_y.sj2[i].gon_cha*10.01),zoj));
				strcat(d_1,ii_bin(p_y.sj2[i].stop_time,zoj));	//胶状
				ad_d(d_1);
				bcom1(d_1);
			}
		}
		head_t(d_1,"WW");
		head_d(d_1,"D",440,5);
		strcat(d_1,ii_bin(p_y.base.max_wen_du,zoj));
		strcat(d_1,ii_bin(p_y.base.min_wen_du,zoj));
		strcat(d_1,ii_bin(p_y.base.pai_liao_time,zoj));
		strcat(d_1,ii_bin(p_y.base.pai_liao_ask,zoj));
		strcat(d_1,ii_bin((int)(p_y.base.weight*100),zoj));
		ad_d(d_1);
		bcom1(d_1);
		head_t(d_1,"WW");
		head_d(d_1,"D",445,5);
		strcat(d_1,ii_bin(0,zoj));
		strcat(d_1,ii_bin(p_y.base.temp1,zoj));
		strcat(d_1,ii_bin(p_y.pei_max,zoj));
		strcat(d_1,ii_bin(run_table_do.total_che,zoj));
		strcat(d_1,ii_bin(p_y.huan[0].speed,zoj));
		ad_d(d_1);
		bcom1(d_1);
		for(i=0;i<p_y.pei_max;i++)
		{
		j=450+(i*10);
		head_t(d_1,"WW");
		head_d(d_1,"D",j,5);
		strcat(d_1,ii_bin(p_y.huan[i].dou_jiao_time,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].dou_tan_time,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].dou_you_time,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].dou_xiao_time,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].st,zoj));
		ad_d(d_1);
		bcom1(d_1);
		head_t(d_1,"WW");
		head_d(d_1,"D",j+5,5);
		strcat(d_1,ii_bin(p_y.huan[i].pre,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].lian_time,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].tem,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].neg,zoj));
		strcat(d_1,ii_bin(p_y.huan[i].ctr,zoj));
		ad_d(d_1);
		bcom1(d_1);
		}
		cs_flag=1;
		write_run(&run_table_do,-1);
		
	 }
	 Sleep(550);

	 i=0;jt=1;
		 time(&tim_3);
		 run_table_do.all_che_time=1;
		 while(i<p_y.pei_max)
		 {
			 time(&tim_1);
			 run_table_do.duan_time=1;
			 do{
				 time(&tim_2);
				 tim1=difftime(tim_2,tim_1);
			     tim2=difftime(tim_2,tim_3);
				 jt=jt+2;
				 k=jt+5;
				 write_power(i+1,jt,k,1);	//写温度、能量进行图表统计				 
				 sprintf(temp,"%3d -- %3d",tim1,p_y.huan[i].lian_time);
				 TextOut(hdc,120,50,temp,strlen(temp));
				 run_table_do.duan_time=tim1;
				 run_table_do.all_che_time=tim2;
				 write_run(&run_table_do,-1);
				 Sleep(550);
			 }while((int)tim1<p_y.huan[i].lian_time);
				jt=0;
			 	Duan_Data.duan_hao=i+1;
				Duan_Data.get_hun_time=(int)tim1;
				Duan_Data.get_you_time=p_y.huan[i].dou_you_time;
				Duan_Data.get_you2_time=p_y.huan[i].dou_you2_time;
				Duan_Data.get_jiao_time=p_y.huan[i].dou_jiao_time;
				Duan_Data.get_tan_time=p_y.huan[i].dou_tan_time;
						
				Duan_Data.next_tempro=jt;
				Duan_Data.next_power=k;
							
				lstrcpy(Duan_Data.duan_begin_time,"11:12:12");
				lstrcpy(Duan_Data.duan_end_time,"11:22:12");

				write_make(Duan_Data);
				sprintf(temp,"%2d 段结束",i+1);
			   x=strlen(temp);
			   TextOut(hdc,220,50,temp,x);
			 i++;
		 }
			write_power(i+1,jt,k,-1);	// 一车结束，结束能量、温度统计
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
	            // 	Make_Temp.th1[i].get_weight=30.0f+i;
					Make_Temp.th1[i].get_weight=p_y.th1[i].weight+2;
			for(i=0;i<Make_Temp.th2_sum;i++)//测试碳黑2生产数据
       	        // 	Make_Temp.th2[i].get_weight=20.0f+i;
					Make_Temp.th2[i].get_weight=p_y.th2[i].weight+2;

			for(i=0;i<Make_Temp.yz1_sum;i++)//测试油料一1生产数据
              	// 	Make_Temp.yz1[i].get_weight=40.0f+i;
					Make_Temp.yz1[i].get_weight=p_y.yz1[i].weight+2;
                
			Make_Temp.get_weight=100.0f; //小料实际重量
			Make_Temp.huan_time=(int)tim2;  //混炼总时间
			Make_Temp.pai_tempro=102;  //排料温度
			Make_Temp.get_pai_time=103;//排料时间
			Make_Temp.all_power=104.0f;  //消耗总能量
			write_che(Make_Temp);	//写一车统计数据
			run_table_do.all_che_time=-1;
			write_run(&run_table_do,-1);
			sprintf(temp,"%2d 车结束",run_table_do.now_che);
			x=strlen(temp);
			TextOut(hdc,320,50,temp,x);
			run_table_do.now_che++;

/**************************
	 add by zgh 4.1
	 
**************************/
		   write_run(&run_table_do,-1);		
		   Sleep(10);
		   read_run(&run_table_do,&run_flag);		// add by zgh 3.8
		   if (run_flag==3)
		   {
			   sprintf(temp,"在线修改%d车结束",run_table_do.now_che);
			   x=strlen(temp);
			   TextOut(hdc,30,80,temp,x);
			   if(run_table_do.now_che>run_table_do.total_che)
			   {
				   sprintf(temp,"在线修改%d车非法",run_table_do.total_che);
				   x=strlen(temp);
				   TextOut(hdc,30,50,temp,x);
					goto exit_rundo;
			   }	
			   else
			   		cs_flag=0;

				write_run(&run_table_do,4);		
		   }
/**************************
	 up add by zgh 4.1
	 
**************************/
		//	now_che++;
		 sprintf(temp,"%d车运行结束",run_table_do.now_che);
		 x=strlen(temp);
		TextOut(hdc,30,50,temp,x);
		
	 }
exit_rundo:
	 write_run(&run_table_do,0);
	 cs_flag=0;
	 sprintf(temp,"运行结束");
	 x=strlen(temp);
	 TextOut(hdc,30,50,temp,x);
	 Sleep(550);

 }
    return FALSE;

}


LRESULT WINAPI FourThread (int x)
{
 int i,j,k;
 char temp[20];
 HDC hdc;
	hdc=GetDC(hwnd);
	 do
	 {
	  Sleep(1100);
     }while(cs_flag==0);
   ReleaseDC(hwnd,hdc);
//   write_multi_screen(200,60,temp);
	
   //测试一车
/*  for(i=0;i<100;i++)
   {
		write_power(1,i+30,(i+30)%40,1);  //write tempro
		write_power(2,i+40,(i+50)%60,1);  //write tempro
		write_power(3,i+20,(i+20)%40,1);  //write tempro
   }
	memcpy(Make_Temp.th1,p_y.th1,sizeof(p_y.th1));			   
	memcpy(Make_Temp.th2,p_y.th2,sizeof(p_y.th2));			   
	memcpy(Make_Temp.yz1,p_y.yz1,sizeof(p_y.yz1));			   
	memcpy(Make_Temp.sj1,p_y.sj1,sizeof(p_y.sj1));			   
	Make_Temp.th1_sum=p_y.th1_sum;
	Make_Temp.th2_sum=p_y.th2_sum;
	Make_Temp.yz1_sum=p_y.yz1_sum;
	Make_Temp.sj1_sum=p_y.sj1_sum;
	for(i=0;i<Make_Temp.th1_sum;i++)//测试碳黑1生产数据
    	Make_Temp.th1[i].get_weight=p_y.th1[i].weight;
                
	for(i=0;i<Make_Temp.th2_sum;i++)//测试碳黑2生产数据
       	Make_Temp.th2[i].get_weight=p_y.th2[i].weight;
                
	for(i=0;i<Make_Temp.yz1_sum;i++)//测试油料1生产数据
    	Make_Temp.yz1[i].get_weight=p_y.yz1[i].weight;

	for(i=0;i<Make_Temp.sj1_sum;i++)//测试胶料1生产数据
    	Make_Temp.sj1[i].get_weight=p_y.sj1[i].weight;

		Make_Temp.get_weight=12; //小料实际重量
		Make_Temp.huan_time=10;  //混炼时间
		Make_Temp.pai_tempro=110;  //排料温度
		Make_Temp.all_power=40;  //消耗总能量
		write_che(Make_Temp);*/
	 do
	 {
	  Sleep(1100);
     }while(cs_flag==1);
/*	int i=0,j,k,t_1=0,t_2=0;
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
	 			head_t(d_1,"WW");
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
				head_t(d_1,"WR");
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
				head_t(d_1,"WR");
				head_d(d_1,"D",1,5);

				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
				bcom4();
				sprintf(temp,"%6.2f",d_f);
                TextOut(hdc,400,250,temp,6);
				 

				head_t(d_1,"WW");
				head_d(d_1,"M",0,1);
				strcat(d_1,h16_16a(0xfae7,zoj));
				ad_d(d_1);
   //             TextOut(hdc,10,80,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"WR");
				head_d(d_1,"M",0,1);
				ad_d(d_1);
 //               TextOut(hdc,10,150,d_1,strlen(d_1));
				bcom1(d_1);
				bcom2();
				sprintf(temp,"%4x",d_d[1]);
                TextOut(hdc,300,40,temp,4);
				head_t(d_1,"BW");
				i=800;if(j == 1) j=0; else j=1;
				head_d(d_1,"M",i,16);
				for(i=0;i<16;i++)
				{if(j==0) j=1; else j=0;
				strcat(d_1,i_bin(j,zoj));}
				ad_d(d_1);
   //             TextOut(hdc,10,250,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"WR");
				head_d(d_1,"M",800,1);
				ad_d(d_1);
//                TextOut(hdc,10,200,d_1,strlen(d_1));
				bcom1(d_1);
				bcom2();
				k=d_d[1];
				sprintf(temp,"%4x",d_d[1]);
                TextOut(hdc,300,120,temp,4);


				head_t(d_1,"WW");
				head_h(d_1,"Y",0xe0,1);
				strcat(d_1,h16_16a(t_2,zoj));
				ad_d(d_1);
     //           TextOut(hdc,10,80,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"WR");
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
			  Sleep(1100);
            return FALSE;
}

LRESULT WINAPI FiveThread (int x)
{
	return 1;
}
//test write_power   write_elec_input  
//                   write_elec_output
struct  elect_send_type     elec_input;  //elec
struct  elect_send_type     elec_output;  //elec


LRESULT WINAPI Read_Elec_Thread (int x)
{
		char temp[380];
		struct send_table  my_send;
	 do{
       if (run_flag==STOP_DOING)
       {			   	
		   	HDC  hdc;
			int x,i,j;
		    init_elec_input(&elec_input);
			init_elec_output(&elec_output);
		

			   hdc=GetDC(hwnd);
			   if(elec_output.flag==2)
				sprintf(temp,"输入信号总长度=%d 输出信号总长度=%d 标志=%2d 状态=调试",
				   elec_input.length,elec_output.length,elec_output.flag);
			   else
				sprintf(temp,"输入信号总长度=%d 输出信号总长度=%d 标志=%2d 状态=正常",
				   elec_input.length,elec_output.length,elec_output.flag);

			   x=strlen(temp);
               TextOut(hdc,20,1,temp,x);				
			    read_send_table(1,1,&my_send);
					x=Read_VbKey();
				if(x==0)				
					sprintf(temp,"碳黑Max=%5.2f,min=%5.2f,AD_mx=%4d,AD_mi=%4d,Port=%2d,Math=%2d,没读入键盘",
							my_send.max,my_send.min,my_send.AD_max,my_send.AD_min,my_send.port,my_send.mathine);
				
				else
					sprintf(temp,"碳黑Max=%5.2f,min=%5.2f,AD_mx=%4d,AD_mi=%4d,Port=%2d,Math=%2d,读入键盘=%4d",
							my_send.max,my_send.min,my_send.AD_max,my_send.AD_min,my_send.port,my_send.mathine,x);
				
			   i=strlen(temp);
               TextOut(hdc,20,20,temp,i);
			   x=40;j=1110;
			   for(i=0;i<elec_input.length;i++)
			   {
				   if(elec_input.note_name[i]==j)
				   {
				   sprintf(temp,"节点名=%5d,数据=%2d",elec_input.note_name[i],elec_input.data[i]);
	               TextOut(hdc,20,x,temp,strlen(temp));
				   x=x+20;
				   j++;
				   }
			   }

				x=40;j=1500;
			   for(i=0;i<elec_output.length;i++)
			   {
	//			   if(elec_input.note_name[i])
	//			   {
				   sprintf(temp,"节点名=%5d,数据=%2d",elec_output.note_name[i],elec_output.data[i]);
	               TextOut(hdc,320,x,temp,strlen(temp));
				   x=x+20;
				   j++;
	//			   }
			   }

			   ReleaseDC(hwnd,hdc);
			   write_multi_screen(120,60,temp);
        }
	   	Sleep(800);
	 }while(1);
        return FALSE;
}


LRESULT WINAPI Read_Run_Thread (int x)
{

	char temp[480];
	int  elec_number;
	static struct	elec_in_table my_elec[41];

	int i,j;
    do{
       static int doing_che=0;
       
       char   *statu[6]={"停止", "运行","模拟"};

       if (run_flag!=STOP_DOING)
       {			   	
		   	HDC  hdc;
			int x;
			   hdc=GetDC(hwnd);

/*			   sprintf(temp,"总车数=%2d,班号=%2d,机号=%2d,配方编号=%s,配方名称=%s,总段数=%2d,当前车=%2d",
							run_table_do.total_che,run_table_do.ban,run_table_do.mathine,run_table_do._number,\
							run_table_do.name,run_table_do.all_duan,run_table_do.now_che);
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

			   for(i=0;i<p_y.yt1_sum;i++)
			   {
			   sprintf(temp,"油二1斗号=%s,次数=%2d,油投1重=%5.2f,快慢=%5.2f,点动=%5.2f,公差=%5.2f,控时=%3.2f,稳时=%2d",
						p_y.yt1[i].dou,p_y.yt1[i].add_time,p_y.yt1[i].weight,p_y.yt1[i].fast_do,\
						p_y.yt1[i].drop_do,p_y.yt1[i].gon_cha,p_y.yt1[i].control_time,p_y.yt1[i].stop_time);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }

			   for(i=0;i<p_y.sj1_sum;i++)
			   {
			   sprintf(temp,"次数=%2d,胶投1重=%5.2f,公差=%5.2f,胶状=%2d",
						p_y.sj1[i].add_time,p_y.sj1[i].weight,p_y.sj1[i].gon_cha,p_y.th1[i].stop_time);
			   x=strlen(temp);j=j+20;
			   TextOut(hdc,20,j,temp,x);
			   }
			   
			   for(i=0;i<p_y.pei_max;i++)
			   {
               sprintf(temp,"投胶=%2d,投炭=%2d,投油=%2d,投小=%2d,转速=%2d,压力=%2d,段时=%2d,温度=%2d,能量=%6.3f,关系=%2d",
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

				elec_number=30;
				Get_Elec_Check_Table(my_elec,&elec_number);
				for(x=0;x<elec_number;x++)
				{
					   sprintf(temp,"name=%d str=%s run=%d,run_set=%d, check=%d, check_set=%d use=%d,time=%d",
						   my_elec[x].name,my_elec[x].str,my_elec[x].run,my_elec[x].run_set,
						   my_elec[x].check,my_elec[x].check_set,my_elec[x].use_flag,my_elec[x].check_time);						
						j=j+20;
						TextOut(hdc,20,j,temp,strlen(temp));				
				}
*/
			   ReleaseDC(hwnd,hdc);
			   write_multi_screen(200,60,temp);
        }
	   Sleep(1000);
    }while(1);
       return FALSE;
}

// test dll
LRESULT WINAPI Write_Elec_Thread (int y)
{
    int k;
    int data=0;
    do{
         if (data==0) data=1;
         else   data=0;
      /*    for(k=0;k<=elec_input.length;k++)
                        elec_input.data[k]=data;
          write_elec_input(&elec_input);

          for(k=0;k<=elec_output.length;k++)
                        elec_output.data[k]=data;
          write_elec_output(&elec_output);
	*/
		  Set_Turn(2482,data);Set_Turn(2483,data);
		  Set_Light(2264,1);Set_Light(2265,1);
          Sleep(1600);
    }while(1);

    return FALSE;
}




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


LONG _EXPORT FAR PASCAL  ThreadWndProc(HWND hwnd,UINT msg,
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

         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) Read_Run_Thread,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;

         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) OneThread ,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;


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

         if (!(hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) FiveThread ,
                            (LPVOID)i,
                             0, &dwThreadID)))
                  MessageBeep(0);		    
            chdthd[total_thread++]=hThread;


        break;
    case WM_COMMAND:
        switch( LOWORD( wparam ) ) 
		{         
        case MENU_ABOUT:
            proc = MakeProcInstance( AboutDlgProc, MyInstance );
            DialogBox( MyInstance,"AboutBox", hwnd, proc );
            FreeProcInstance( proc );
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
		}//end case
		break;
    case WM_DESTROY:
         x=sio_close(TEST_PORT);
		for(i=0;i<total_thread;i++)		 
		{			
			TerminateThread(chdthd[i],1);
		}
		total_thread=0;
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
 int j,k,l;
 char p2[7];
 strcat(p,p1);
 j=m+10000;
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
 int j,k,l;
 char p2[7];
 strcat(p,p1);
 j=m+0x10000;
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
 for(i=1;i<j;i++) p1=p1+(*(p+i));
 if((p1&0xf0)>0x90) *(p+i)=0x37+((p1&0xf0)>>4);
 else *(p+i)=0x30+((p1&0xf0)>>4);
 if((p1&0xf)>0x9) *(p+i+1)=0x37+(p1&0xf);
 else *(p+i+1)=0x30+(p1&0xf);
 *(p+i+2)=0;
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
//	sio_putch(2,0x04);
//	Sleep(20);
	sio_putch(2,0x0c);
	Sleep(10);
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
 HDC  hdc;
// sio_flush(2,0);	
// Sleep(50);
// sio_putch(2,0x0c);
// Sleep(60);
 hdc=GetDC(hwnd);
 for(i=0;i<8;i++)
     temp[i]=' ';
    temp[i]='\0';

 i=0;
 do{
 //   sio_read(2,temp,1);
//	temp[i]=temp[0]+0x30;
//	temp[i]=j+0x30;
//    TextOut(hdc,10,100,temp,9);
	j=sio_getch(2);
//	Sleep(2);
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
 ReleaseDC(hwnd,hdc);
 //return k;
}
void bcom3()
{
 int i,j,k;
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
 ReleaseDC(hwnd,hdc);
 //return k;
}

void bcom4()
{
 int i,j,k;
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

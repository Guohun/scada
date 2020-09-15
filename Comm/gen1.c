/*
 *%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 *%                                                                        %
 *%     Copyright (C) 1991,1994 by WATCOM International Inc.               %
 *%     All rights reserved.                                               %
 *%                                                                        %
 *%     Permission is granted to anyone to use this example program for    %
 *%     any purpose on any computer system, subject to the following       %
 *%     restrictions:                                                      %
 *%                                                                        %
 *%     1. This example is provided on an "as is" basis, without warranty. %
 *%        You indemnify, hold harmless and defend WATCOM from and against %
 *%        any claims or lawsuits, including attorney's, that arise or     %
 *%        result from the use or distribution of this example, or any     %
 *%        modification thereof.                                           %
 *%                                                                        %
 *%     2. You may not remove, alter or suppress this notice from this     %
 *%        example program or any modification thereof.                    %
 *%                                                                        %
 *%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 *
 * GEN1.C
 *
 * Generic windows program, demonstrates basic message handling
 *
 */
#include <windows.h>
#include <stdio.h>
#include <malloc.h>
#include "c:\ljxt\genwin.h"
#include<string.h>
#include<conio.h>
#include"Api232-c.h"
#include"m_kl.h"

HANDLE   hThread;
HANDLE    chdthd[10];
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
char d_3[6],d_4[6],m_d[10][6];
int  d_d[10];
float d_f;

LRESULT CALLBACK  ThreadWndProc(HWND hwnd,UINT msg,
                                     UINT wparam, LONG lparam );
HWND hwnd;  //main hwnd
//int  _EXPORT  sio_getch(int),sio_read(int,char,int);
int a1=0,a2=0,a3=0;
typedef struct 
{
        HANDLE          hInstance;
        HWND            hwndFirst;
}THREADDATA;

typedef THREADDATA * PTHREADDATA;


HANDLE          MyInstance;
static char     GenericClass[32]="GenericClass";

static BOOL FirstInstance( HANDLE );
static BOOL AnyInstance( HANDLE, int, LPSTR );

long _EXPORT FAR PASCAL WindowProc( HWND, unsigned, UINT, LONG );

/*
 * WinMain - initialization, message loop
 */
int PASCAL WinMain( HANDLE this_inst, HANDLE prev_inst, LPSTR cmdline,
                    int cmdshow )
{
    MSG         msg;

    MyInstance = this_inst;
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

    return( msg.wParam );

} /* WinMain */


static          char szClassName [21] ="SCONDTHREAD";
static          char szWindowTite [21] = "Scond Thread";
LRESULT WINAPI OneThread (int x)
{
	int i=0,j,k;
	float f1;
	char *s1,*s2;
	unsigned c1;
	char temp[20];
	HDC  hdc;
              while (TRUE)
               {
                hdc=GetDC(hwnd);
               do
			   {
			    i=sio_getch(3);
			   }while(i != 0x02);
			   sio_read(3,s1,9);
			   c1=*(s1+0)&0x007;
			   for(i=0;i<7;i++)
				   *(s2+i)=*(s1+i+3);
			   k=atoi(s2);
			   switch(c1)
			   {
			   case 0x01: f1=k*10.0;break;
			   case 0x03: f1=k/10.0;break;
			   case 0x04: f1=k/100.0;break;
			   case 0x05: f1=k/1000.0;break;
			   case 0x06: f1=k/10000.0;break;
			   };
			   if((*(s1+1)&0x02)!=0) f1=-f1;
				sprintf(temp,"%8.4f",f1);
                TextOut(hdc,10,100,temp,8);
                ReleaseDC(hwnd,hdc);
                 Sleep(1000);

               };
            return FALSE;
}

LRESULT WINAPI SecondThread (int x)
{
    char  temp[40],temp1[40];
    int i=0,j,k;
	float f1;
	char s1[40],s2[40];
	unsigned c1;

	HDC  hdc;
/*          x=sio_close(2);
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
                 temp1[i]='\0'; */
              while (TRUE)
               {
                hdc=GetDC(hwnd);
  	/*		    sio_flush(2,0);	
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
   				   sio_flush(2,0); */
                   Sleep(200);

               };
                ReleaseDC(hwnd,hdc);
             /*   x=sio_close(2);
                  if (x<0)
                  {
                       MessageBox(hwnd,"close  error","open",0);
                        return  0;
                 }*/

}

LRESULT WINAPI ThreeThread (int x)
{
	int i=0,j,k;
	float f1;
	char *s1,*s2;
	unsigned c1;
	char temp[20];
	HDC  hdc;
    
              while (TRUE)
               {
                hdc=GetDC(hwnd);
               do
			   {
			    i=sio_getch(5);
			   }while(i != 0x02);
			   sio_read(5,s1,9);
			   c1=*(s1+0)&0x007;
			   for(i=0;i<7;i++)
				   *(s2+i)=*(s1+i+3);
			   k=atoi(s2);
			   switch(c1)
			   {
			   case 0x01: f1=k*10.0;break;
			   case 0x03: f1=k/10.0;break;
			   case 0x04: f1=k/100.0;break;
			   case 0x05: f1=k/1000.0;break;
			   case 0x06: f1=k/10000.0;break;
			   };
			   if((*(s1+1)&0x02)!=0) f1=-f1;
				sprintf(temp,"%8.4f",f1);
                TextOut(hdc,10,100,temp,8);
                ReleaseDC(hwnd,hdc);
                 Sleep(1000);

               };
            return FALSE;

}

LRESULT WINAPI FourThread (int x)
{
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
	/*			head_t(d_1,"WW");
				i=0;
				head_d(d_1,"D",10,6);
				for(i=t_1;i<t_1+6;i++)
				strcat(d_1,ii_bin(i,zoj));
				ad_d(d_1);
                TextOut(hdc,10,200,d_1,strlen(d_1));
				bcom1(d_1);
	//		    sio_flush(2,2);	
    //            Sleep(2);
	//			for(i=0;i<5;i++)
	//			{
				head_t(d_1,"WR");
				head_d(d_1,"D",0,6);
				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
	//			sio_write(2,d_1,strlen(d_1));
				bcom3();
				for(i=1;i<7;i++)
				{
				sprintf(temp,"%4d",d_d[i]);
                TextOut(hdc,400,(20+i*20),temp,4);
				} */
				head_t(d_1,"WR");
				head_d(d_1,"D",1,5);
				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
				bcom4();
				sprintf(temp,"%6.2f",d_f);
                TextOut(hdc,400,260,temp,6);
				 

				head_t(d_1,"WW");
				head_d(d_1,"M",0,1);
				strcat(d_1,h16_16a(0xfae7,zoj));
				ad_d(d_1);
   //             TextOut(hdc,10,80,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"WR");
				head_d(d_1,"M",0,1);
				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
				bcom2();
				sprintf(temp,"%4x",d_d[1]);
                TextOut(hdc,300,40,temp,4);
				head_t(d_1,"BW");
				i=800;j=1;
				head_d(d_1,"M",i,25);
				for(i=0;i<16;i++)
				{if(j==0) j=1; else j=0;
				strcat(d_1,i_bin(j,zoj));}
				ad_d(d_1);
				bcom1(d_1);

				head_t(d_1,"WW");
				head_h(d_1,"Y",0xc0,1);
				strcat(d_1,h16_16a(t_2,zoj));
				ad_d(d_1);
                TextOut(hdc,10,80,d_1,strlen(d_1));
    			bcom1(d_1);
				head_t(d_1,"WR");
				head_h(d_1,"Y",0xc0,1);
				ad_d(d_1);
    //            TextOut(hdc,10,300,d_1,strlen(d_1));
				bcom1(d_1);
				bcom2();
				sprintf(temp,"%4x",d_d[1]);
                TextOut(hdc,300,80,temp,4);

				t_1=t_1+5;
				t_2=t_2+0x421;
				if(t_1>2000) t_1=0;
				if(t_2>0xfff0) t_2=0;

	//			sio_write(2,d_1,strlen(d_1));
			    sio_flush(2,1);	

                ReleaseDC(hwnd,hdc);
                Sleep(55);

               };
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
//    wc.hIcon = LoadIcon( this_inst, "GenericIcon" );
    wc.hCursor = LoadCursor( NULL, IDC_ARROW );
    wc.hIcon = LoadIcon (NULL,IDI_APPLICATION);
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
        "WATCOM Generic Kind Of Application",   /* caption */
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



LONG _EXPORT FAR PASCAL WindowProc( HWND hwnd, unsigned msg,
                                     UINT wparam, LONG lparam )
{
    FARPROC     proc;
    extra_data  *edata_ptr;
    char        buff[128];
    int i;

    DWORD    dwThreadID;
    THREADDATA  trdata;

    switch( msg ) {

    case WM_COMMAND:
        switch( LOWORD( wparam ) ) {
		case MENU_KILL:
		for(i=0;i<4;i++)		 
		{
                
                 TerminateThread(chdthd[i],1);
		}
		 break;
         
        case MENU_ABOUT:
            proc = MakeProcInstance( AboutDlgProc, MyInstance );
            DialogBox( MyInstance,"AboutBox", hwnd, proc );
            FreeProcInstance( proc );
            break;

        case MENU_CMDSTR:
		
          //creating One thread
		 
		 
          hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) OneThread,
                            (LPVOID)i,
                             0, &dwThreadID);
                     		 MessageBeep(0);
			chdthd[0]=hThread;
		  //creating second thread
		 
		 
          hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) SecondThread,
                            (LPVOID)i,
                             0, &dwThreadID);
                     		 MessageBeep(0);
			chdthd[1]=hThread;

          hThread = CreateThread(NULL, 0,
                            (LPTHREAD_START_ROUTINE) ThreeThread,
                            (LPVOID)i,
                             0, &dwThreadID);
                     		 MessageBeep(0);
			chdthd[2]=hThread;
         
			if(!(hThread = CreateThread(NULL, 0,
                (LPTHREAD_START_ROUTINE) FourThread,
                (LPVOID)i,
                0, &dwThreadID)))
              MessageBeep(0);
              chdthd[3]=hThread;
             }
        break;

    case WM_DESTROY:
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
	Sleep(2);
   }while(d_2[i-1] != 0x03);
 d_2[i-1]=0;
// TextOut(hdc,10,250,d_2,strlen(d_2));
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
 Sleep(2);
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
	Sleep(2);
   }while(d_2[i-1] != 0x03);
 d_2[i-1]=0;
 TextOut(hdc,10,250,d_2,strlen(d_2));
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
 Sleep(2);
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
	Sleep(2);
   }while(d_2[i-1] != 0x03);
 d_2[i-1]=0;
 // TextOut(hdc,10,250,d_2,strlen(d_2));
 d_3[0]=d_2[19];
 d_3[1]=d_2[17];
 d_3[2]=d_2[23];
 d_3[3]=d_2[21];
 d_3[4]='/0';
 d_d[0]=asc_10i(d_3);
 d_f=d_d[0]/100.0;
 sio_flush(2,1);	
 Sleep(2);
 ReleaseDC(hwnd,hdc);
 //return k;
}
/*
void main()
{
 int i,j,k;
 char p1,p2,zoj[10],zo1[10];

 // _bios_serialcom(_COM_INIT,1,0xfa);
 bioscom(0,0xea,1);

 head_t(d_1,"WW");
 i=0;
 head_d(d_1,"D",i,25);
 for(i=200;i<225;i++)
 strcat(d_1,ii_bin(i,zoj));
// strcat(d_1,ii_bin(498,zoj));
// strcat(d_1,h16_16a(789,zoj));
// strcat(d_1,h16_16a(256,zoj));
// strcat(d_1,h16_16a(3689,zoj));
 ad_d(d_1);
 printf(" A=%s \n",d_1);
 bcom1(d_1);
 head_t(d_1,"BW");
 i=800;j=1;
 head_d(d_1,"M",i,25);
 for(i=0;i<25;i++)
 {if(j==0) j=1; else j=0;
 strcat(d_1,i_bin(j,zoj));}
 ad_d(d_1);
 printf(" B=%s \n",d_1);
 bcom1(d_1);
 head_t(d_1,"WR");
 head_d(d_1,"D",0,5);
 ad_d(d_1);
 bcom1(d_1);
 bcom2();
 for(i=1;i<5;i++)
 printf(" k=%d \n",d_d[i]);
} */
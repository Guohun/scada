
/* ---------------- */
/* FileName: svga.c */
/* ---------------- */

#include <windows.h>
#include <conio.h>
  #include "SQL.H"
  #include "SQLext.H"

HBITMAP  hBit;
#include "resource.h"
#include "dll.h"
#include "temp.h"
HINSTANCE   MyInstance;



#pragma     data_seg("sdata")
#define MAX_SCREEN_ROW  8
	int   nMaxX=0;
	int   nMaxY=0;
	int    screen_rear=0;
	int    screen_front=0;
	int    screen_flag=0;
    struct  screen_buff  desk_screen[MAX_SCREEN_ROW]={1,0,0,0,"",1,0,0,0,"",
											  1,0,0,0,"",1,0,0,0,"",
											  1,0,0,0,"",1,0,0,0,"",
											  1,0,0,0,"",1,0,0,0,""
/*											  1,0,0,0,"",1,0,0,0,""
											  1,0,0,0,"",1,0,0,0,"",
											  1,0,0,0,"",1,0,0,0,"",
											  1,0,0,0,"",1,0,0,0,"",
											  1,0,0,0,"",1,0,0,0,"",
											  1,0,0,0,"",1,0,0,0,""											  
*/

	};
		
		init_screen_flag=0;
		HDC   hdc=0;
		HANDLE m_hEventWrite=NULL;// can send  doing
		HDC   memdc=0;
#pragma data_seg()


int WINAPI  Init_Font(int size,char * name)
{
	char   temp[40];
    if  (size==-1)
    {
        FONT_SIZE=0;
        return 1;
    }
	lstrcpy(temp,name);
	if (hNewf!=NULL) DeleteObject(hNewf);
       hNewf=CreateFont(size,0,0,0,FW_NORMAL,0,0,0,ANSI_CHARSET,
			OUT_DEFAULT_PRECIS,CLIP_LH_ANGLES ,
            DEFAULT_QUALITY,DEFAULT_PITCH|FF_DONTCARE,temp);
	if (hNewf==NULL)  
	{
		FONT_SIZE=0;
		return 0;
	}
	FONT_SIZE=size;
	return 1;
}

short WINAPI   VBwrite_multi_screen(void)
{
 
        HFONT old_font;
		int len=0;

		if (!init_screen_flag)
		{		
			m_hEventWrite=CreateEvent(NULL, TRUE, TRUE, "Write_Screen");
			init_screen_flag=1;
		}


//        {
            if (desk_screen[0].Clear_flag==1)
            {
  
//                SetThreadPriority(GetCurrentThread(),THREAD_PRIORITY_TIME_CRITICAL);
			if (screen_flag!=2)
			{
               _outp(0x2c8,0xfe);
               _outp(0x2cb,0xff);
               _outp(0x2c9,~(0x11<<1));    //section       write
			}

//				hdc=GetDC(0);			
				if (screen_flag==2)
					BitBlt(hdc,0,nMaxY/2+10,nMaxX,nMaxY,memdc,0,0,SRCCOPY);
				else
					BitBlt(hdc,0,0,nMaxX,nMaxY,memdc,0,0,SRCCOPY);
				if (screen_flag!=2)
				{
                 _outp(0x2c8,0xfe);
                 _outp(0x2cb,0xff);
                 _outp(0x2c9,~(0x11<<0)); //section resum
				}
//                DeleteDC(memdc);
//				 ReleaseDC(0,hdc);

//                SetThreadPriority(GetCurrentThread(),THREAD_PRIORITY_IDLE);
                desk_screen[0].Clear_flag=0;
                return 1;
            }

//	for(git=0;git<MAX_SCREEN_ROW;git++)			
//	   if (desk_screen[screen_front].change_flag==1)  
		if (screen_front!=screen_rear)		
		{
		if (screen_flag!=2)
		{
//                SetThreadPriority(GetCurrentThread(),THREAD_PRIORITY_TIME_CRITICAL);
			   if  (WaitForSingleObject(m_hEventWrite, 0) == WAIT_TIMEOUT)  return 0;
					   ResetEvent(m_hEventWrite);
						   _outp(0x2c8,0xfe);
                           _outp(0x2cb,0xff);
                           _outp(0x2c9,~(0x11<<1));    //section       write
					{
//					hdc=GetDC(0);
                   if(FONT_SIZE!=0)
					old_font=(HFONT )SelectObject(hdc,hNewf);
					for(;screen_front!=screen_rear;)					
					{
                                    screen_front=(screen_front+1) %MAX_SCREEN_ROW;
                                    len=strlen(desk_screen[screen_front].str);
                                    TextOut(hdc,desk_screen[screen_front].col,desk_screen[screen_front].row,desk_screen[screen_front].str,len);
                    }
					screen_front=0;
					screen_rear=0;
				}
				

				if(FONT_SIZE!=0)
						SelectObject(hdc,old_font);
//				ReleaseDC(0,hdc);
//              InvalidateRect(GetDesktopWindow(),NULL,1);
                             _outp(0x2c8,0xfe);
                             _outp(0x2cb,0xff);
                             _outp(0x2c9,~(0x11<<0)); //section resum
					SetEvent(m_hEventWrite);
                //             SetThreadPriority(GetCurrentThread(),THREAD_PRIORITY_IDLE);
		}
		else//end screen_flag
		{

			   if  (WaitForSingleObject(m_hEventWrite, 0) == WAIT_TIMEOUT)  return 0;
			   ResetEvent(m_hEventWrite);

//				hdc=GetDC(0);
               if(FONT_SIZE!=0)
				old_font=(HFONT )SelectObject(hdc,hNewf);
			    
					for(;screen_front!=screen_rear;)					
					{
                                    screen_front=(screen_front+1) %MAX_SCREEN_ROW;
                                    len=strlen(desk_screen[screen_front].str);
                                    TextOut(hdc,desk_screen[screen_front].col,desk_screen[screen_front].row+nMaxY/2+10,desk_screen[screen_front].str,len);
					}
					screen_front=0;
					screen_rear=0;
				}
				if(FONT_SIZE!=0)
						SelectObject(hdc,old_font);
//				ReleaseDC(0,hdc);
				SetEvent(m_hEventWrite);


//			  desk_screen[screen_front].change_flag=0;
        }//end if	
	//end for dit
//        
        return len;
}
int WINAPI Clear_multi_screen(void)
{
 
	desk_screen[0].Clear_flag=1;
	return 1;
}

short WINAPI Set_multi_screen(short flag)
{	
	HBRUSH  hBrush;
	screen_flag=flag;
    nMaxX=GetSystemMetrics(SM_CXSCREEN);
    nMaxY=GetSystemMetrics(SM_CYSCREEN);
    hdc=GetDC(0);
                memdc=CreateCompatibleDC(hdc);
                hBit=CreateCompatibleBitmap(hdc,nMaxX,nMaxY);

                SelectObject(memdc,hBit);
//                hBit=LoadBitmap(NULL,"ZGH_BMP");
                //hBit=CreateBitmap(16,16,1,1,(LPSTR)str);

                hBrush=GetStockObject(WHITE_BRUSH);
                SelectObject(memdc,hBrush);
	PatBlt(memdc,0,0,nMaxX,nMaxY,PATCOPY);
//		ReleaseDC(0,hdc);
	return flag;
}

int WINAPI write_multi_screen(int row,int col,char *str)
{
	int i=0;
	
	char *temp=str;
//	git=(row/MAX_SCREEN_ROW)%MAX_SCREEN_ROW;
		if (!init_screen_flag)
		{
			m_hEventWrite=CreateEvent(NULL, TRUE, TRUE, "Write_Screen");
			init_screen_flag=1;
		}
	if  (WaitForSingleObject(m_hEventWrite, 0) == WAIT_TIMEOUT)  return 0;
	ResetEvent(m_hEventWrite);
	{		
		if ((screen_rear+1)==MAX_SCREEN_ROW)	 
			{
				SetEvent(m_hEventWrite);
				return 0;
			}
		else
			screen_rear=(screen_rear+1) % MAX_SCREEN_ROW;	
	}
	SetEvent(m_hEventWrite);

	desk_screen[screen_rear].row=row;
	desk_screen[screen_rear].col=col;	
	for(i=0;i<99&&*temp!='\0';i++,temp++)
			desk_screen[screen_rear].str[i]=*temp;
	desk_screen[screen_rear].str[i]='\0';
	desk_screen[screen_rear].change_flag=1;	

        return i;
}
